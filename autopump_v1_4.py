import os
import sys

# Set DPI awareness before Qt imports to avoid warnings
if sys.platform == "win32":
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"
import io
import sys
import threading
import time
from typing import Any

import keyboard
import pythoncom
import serial
from PyQt6.QtCore import QObject, QRect, Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QAction, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMenu,
    QPlainTextEdit,
    QPushButton,
    QSpinBox,
    QSystemTrayIcon,
    QVBoxLayout,
    QWidget,
)
from pywinauto import Desktop
import pywinauto


class OutputRedirect(io.StringIO):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def write(self, text):
        self.callback(text)


class PumpController(QObject):
    # Signal for thread-safe logging
    log_message = pyqtSignal(str)
    # Signal to unlock GUI inputs when manual sample entry is needed
    unlock_for_input = pyqtSignal()
    # Signal to relock GUI inputs after sample is submitted
    relock_inputs = pyqtSignal()
    def __init__(self):
        super().__init__()
        self.SERIAL_PORT = "COM5"
        self.runtime_sec = 60  # Default fallback (used by Wash)

        # Replicate tracking
        self.replicate_durations = [
            5,
            5,
            5,
        ]  # Default durations for 3 replicates         
        self.current_replicate = 0  # Track which replicate we're on (0-2)         
        self.in_sample_sequence = False  # Track if we're in an active sample

        # Repeat sample tracking
        self.base_sample_name = ""  # User-supplied base sample name         
        self.repeat_count = 1  # Number of times to repeat the sample

        self.current_repeat = 0  # Current repeat iteration (0-based)
        self.total_repeats_completed = 0  # Total count of repeats done
        self.last_sample = False  # Whether to submit "Yes" for last sample field
        self.pending_window = None  # Start sample window awaiting manual submission

        self.window_name = "Operator Request"
        self.automation_active = False
        self.f8_listener_thread = None
        self.window_monitor_thread = None

    def log(self, message):
        """Thread-safe logging method"""
        self.log_message.emit(message)

    def get_current_duration(self):
        """Get the duration for the current replicate"""
        if self.in_sample_sequence and 0 <= self.current_replicate < 3:
            return self.replicate_durations[self.current_replicate]
        return self.runtime_sec  # Fallback

    def get_current_sample_name(self):
        """Generate the current sample name with letter suffix"""
        if not self.base_sample_name:
            return ""

        # Add letter suffix (A, B, C, etc.)
        if self.repeat_count > 1:    
            letter_index = self.total_repeats_completed % 26
            suffix = chr(ord("A") + letter_index)
            return f"{self.base_sample_name} {suffix}"
        else:
            return f"{self.base_sample_name}"

    def is_last_repeat(self):
        """Check if this is the last repeat"""
        return self.current_repeat >= self.repeat_count - 1

    def pump_sample(self):
        """Execute a pump sample cycle via serial communication"""
        duration = self.get_current_duration()
        try:
            with serial.Serial(self.SERIAL_PORT, 1200, timeout=1) as ser:
                ser.write(f"START:{duration}\n".encode())
                self.log(
                    f"Command sent: START:{duration} (Replicate {self.current_replicate + 1}/3)"
                )
                while True:
                    line = ser.readline().decode().strip()
                    if line:
                        self.log(f"Pico says: {line}")
                    if line == "Pump stopped":
                        break
                    time.sleep(0.5)

            # Advance to next replicate after successful pump
            if self.in_sample_sequence:
                self.log(
                        f"Sample {self.get_current_sample_name()} complete (all 3 replicates finished)"
                    )
                self.current_replicate += 1
                if self.current_replicate >= 3:
                    self.in_sample_sequence = False
                    self.current_replicate = 0

        except Exception as e:
            self.log(f"Pump sample error: {e}")

    def click_button_by_text(self, window, button_text):
        """Click a button with specific text"""
        try:
            for ctrl in window.descendants():
                if ctrl.friendly_class_name() == "Button":
                    if button_text.lower() in ctrl.window_text().lower():
                        ctrl.set_focus()
                        ctrl.click_input()
                        self.log(f"Clicked '{button_text}'")
                        return True
            self.log(f"'{button_text}' button not found")
            return False
        except Exception as exc:
            self.log(f"Failed to click {button_text}: {exc}")
            return False

    def fill_start_sample_window(self, window):
        """Fill in the sample name and last sample fields"""
        try:
            sample_name = self.get_current_sample_name()
            is_last = "Yes" if self.last_sample else "No"

            # Collect only Edit controls so indexes are predictable
            edit_boxes = [
                child for child in window.descendants()
                if child.friendly_class_name() == "Edit"
            ]

            try:
                lims_field = edit_boxes[0]
                lims_field.set_focus()
                self.log("Setting sample name field")
                lims_field.set_edit_text(sample_name)
                self.log(f"Filled sample name (LIMS #): {sample_name}")
            except Exception as e:
                self.log(f"Could not find LIMS # field: {e}")
                return False

            try:
                last_sample_field = edit_boxes[1]
                last_sample_field.set_focus()
                self.log("Setting last sample field")
                last_sample_field.set_edit_text(is_last)
                self.log(f"Filled last sample field: {is_last}")
            except Exception as e:
                self.log(f"Could not find last sample field: {e}")
                return False

            return True
        except Exception as exc:
            self.log(f"Error filling start sample window: {exc}")
            return False

    def is_correct_operator_request(self, window):
        """Check if the window contains the correct operator request text"""
        try:
            for ctrl in window.descendants():
                if ctrl.friendly_class_name() == "Edit":
                    text = ctrl.window_text().strip()
                    if "Pull sample through flow cell" in text:
                        return True
            return False
        except Exception as exc:
            self.log(f"Window check error: {exc}")
            return False

    def is_start_sample_window(self, window):
        """Check if this is a start sample window (no continue button, has accept button)"""
        try:
            has_continue = False
            has_accept = False

            for ctrl in window.descendants():
                if ctrl.friendly_class_name() == "Button":
                    button_text = ctrl.window_text().lower()
                    if "continue" in button_text:
                        has_continue = True
                    if "accept" in button_text:
                        has_accept = True

            # Start sample window has accept but no continue button
            return has_accept and not has_continue
        except Exception as exc:
            self.log(f"Start sample window check error: {exc}")
            return False

    def handle_start_sample_window(self, window):
        """Handle the start sample window by filling in info and clicking accept"""
        try:
            # Check if we should handle this window automatically
            if self.current_repeat < self.repeat_count and self.base_sample_name:
                sample_name = self.get_current_sample_name()
                self.log(
                    f"Start sample window detected - Sample {sample_name} (repeat {self.current_repeat + 1}/{self.repeat_count})"
                )

                # Reset replicate counter for new sample
                self.current_replicate = 0
                self.in_sample_sequence = True

                # Fill in the sample information
                if self.fill_start_sample_window(window):
                    time.sleep(1)  # Brief pause for UI to update

                    # Click the Accept button
                    if self.click_button_by_text(window, "Accept"):
                        self.log(f"Sample {sample_name} started successfully")
                        # Increment counters
                        self.current_repeat += 1
                        self.total_repeats_completed += 1

                        # Check if we've completed all repeats
                        if self.current_repeat >= self.repeat_count:
                            self.log(
                                f"All {self.repeat_count} repeat(s) completed. Automation will pause at next sample."
                            )
                    else:
                        self.log("Failed to click Accept button")
                else:
                    self.log("Failed to fill sample information")
            else:
                # No more configured repeats - unlock GUI for manual entry
                self.log(
                    "Start sample window detected - unlocking inputs for next sample entry"
                )
                self.pending_window = window
                self.unlock_for_input.emit()

        except Exception as exc:
            self.log(f"Error handling start sample window: {exc}")

    def submit_sample(self):
        """Called by the Submit Sample button - fills and accepts the pending start sample window"""
        if self.pending_window is None:
            self.log("No pending sample window to submit")
            return
        try:
            window = self.pending_window
            self.pending_window = None

            sample_name = self.get_current_sample_name()
            self.log(f"Submitting sample: {sample_name}")

            # Reset replicate counter for new sample
            self.current_replicate = 0
            self.in_sample_sequence = True

            if self.fill_start_sample_window(window):
                time.sleep(1)
                if self.click_button_by_text(window, "Accept"):
                    self.log(f"Sample {sample_name} submitted successfully")
                    self.current_repeat += 1
                    self.total_repeats_completed += 1
                    # Relock inputs - they stay locked until the next manual unlock
                    self.relock_inputs.emit()
                else:
                    self.log("Failed to click Accept - window may have closed")
                    self.relock_inputs.emit()
            else:
                self.log("Failed to fill sample information")
                self.relock_inputs.emit()
        except Exception as exc:
            self.log(f"Error submitting sample: {exc}")
            self.relock_inputs.emit()

    def listen_for_f8(self):
        """Background thread to listen for F8 key presses"""
        cooldown_sec = max(self.replicate_durations) + 1
        last_trigger_time = 0
        while self.automation_active:
            try:
                keyboard.wait("f8")
                if not self.automation_active:
                    break
                now = time.time()
                if now - last_trigger_time >= cooldown_sec:
                    self.log("F8 pressed - starting manual injection")
                    last_trigger_time = now
                    self.pump_sample()
                else:
                    self.log(
                        f"F8 pressed too soon. Wait {cooldown_sec - (now - last_trigger_time):.1f} more seconds."
                    )
            except Exception as e:
                self.log(f"F8 listener error: {e}")
                time.sleep(1)

    def monitor_windows(self):
        """Background thread to monitor for operator request windows"""
        pythoncom.CoInitialize()
        try:
            while self.automation_active:
                try:
                    windows = Desktop(backend="win32").windows()
                    for win in windows:
                        if not self.automation_active:
                            return
                        if self.window_name in win.window_text():
                            self.log(f"Found window titled: {win.window_text()}")
                            # Check if this is a start sample window first
                            if self.is_start_sample_window(win):
                                self.handle_start_sample_window(win)
                                time.sleep(2)
                                break

                            # Otherwise, check if it's the correct operator request
                            elif self.is_correct_operator_request(win):
                                self.log(
                                    f"Replicate {self.current_replicate + 1}/3 'Operator Request' detected!"
                                )
                                self.pump_sample()
                                self.click_button_by_text(win, "Continue")
                                time.sleep(3)
                                break
                    time.sleep(1)
                except Exception as e:
                    self.log(f"Window monitor error: {e}")
                    time.sleep(1)
        finally:
            pythoncom.CoUninitialize()

    def start_automation(self):
        """Start the pump automation system"""
        if self.automation_active:
            self.log("Automation already running")
            return
        self.automation_active = True

        # Reset repeat tracking when starting automation
        self.current_repeat = 0
        self.total_repeats_completed = 0

        self.log("Starting pump automation...")
        self.log(
            f"Replicate durations set to: R1={self.replicate_durations[0]} s, R2= {self.replicate_durations[1]} s, R3={self.replicate_durations[2]} s"
        )

        if self.base_sample_name and self.repeat_count > 0:
            self.log(
                f"Repeat samples configured: Base name '{self.base_sample_name}', {self.repeat_count} repeat(s)"
            )
        else:
            self.log("No repeat samples configured - manual sample entry required")
        # Start F8 listener thread
        self.f8_listener_thread = threading.Thread(
            target=self.listen_for_f8, daemon=True
        )
        self.f8_listener_thread.start()
        # Start window monitor thread
        self.window_monitor_thread = threading.Thread(
            target=self.monitor_windows, daemon=True
        )
        self.window_monitor_thread.start()
        self.log(
            "Automation threads started. F8 manual trigger and window monitoring active."
        )

    def stop_automation(self):
        """Stop the pump automation system"""
        if not self.automation_active:
            self.log("Automation already stopped")
            return
        self.log("Stopping pump automation...")
        self.automation_active = False
        # Reset sample tracking
        self.in_sample_sequence = False
        self.current_replicate = 0
        self.pending_window = None
        self.log("Automation stopped. All monitoring threads will terminate.")


class LogWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(
            Qt.WindowType.Tool
            | Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setWindowTitle("Automation Log")
        # Initialize pump controller
        self.pump_controller = PumpController()
        # Connect pump controller's log signal to our GUI update method
        self.pump_controller.log_message.connect(self.append_text)
        self.pump_controller.unlock_for_input.connect(self.unlock_sample_inputs)
        self.pump_controller.relock_inputs.connect(self.lock_sample_inputs)
        # Create layout
        layout = QVBoxLayout()
        # Add repeat sample controls
        repeat_layout = QVBoxLayout()
        repeat_layout.addWidget(QLabel("Repeat Sample Configuration:"))
        # Sample name input
        name_layout = QHBoxLayout()
        name_layout.addWidget(QLabel("Sample Name:"))
        self.sample_name_input = QLineEdit()
        self.sample_name_input.setPlaceholderText("Enter base sample name")
        self.sample_name_input.textChanged.connect(self.update_sample_config)
        name_layout.addWidget(self.sample_name_input)
        repeat_layout.addLayout(name_layout)

        # Last sample checkbox + wash + submit buttons on the same row
        checkbox_button_layout = QHBoxLayout()
        self.last_sample_checkbox = QCheckBox("Last Sample")
        self.last_sample_checkbox.setChecked(False)
        self.last_sample_checkbox.stateChanged.connect(self.update_sample_config)
        checkbox_button_layout.addWidget(self.last_sample_checkbox, stretch=1)
        self.wash_button = QPushButton("Wash")
        self.wash_button.clicked.connect(self.run_wash)
        checkbox_button_layout.addWidget(self.wash_button, stretch=1)
        self.submit_button = QPushButton("Submit Sample")
        self.submit_button.clicked.connect(self.submit_sample)
        self.submit_button.setEnabled(False)  # Only active when inputs are unlocked
        checkbox_button_layout.addWidget(self.submit_button, stretch=1)
        repeat_layout.addLayout(checkbox_button_layout)

        # Repeat count input
        repeat_count_layout = QHBoxLayout()
        repeat_count_layout.addWidget(QLabel("Repeat Count:"))
        self.repeat_count_spinbox = QSpinBox()
        self.repeat_count_spinbox.setRange(0, 100)
        self.repeat_count_spinbox.setValue(1)
        self.repeat_count_spinbox.valueChanged.connect(self.update_sample_config)
        repeat_count_layout.addWidget(self.repeat_count_spinbox)
        repeat_layout.addLayout(repeat_count_layout)

        layout.addLayout(repeat_layout)
        # Add separator
        layout.addWidget(QLabel("─" * 43))
        # Add replicate duration controls
        duration_layout = QVBoxLayout()
        duration_layout.addWidget(QLabel("Replicate Pump Durations (seconds):"))

        # Replicate 1 duration
        r1_layout = QHBoxLayout()
        r1_layout.addWidget(QLabel("Replicate 1:"))
        self.r1_spinbox = QSpinBox()
        self.r1_spinbox.setRange(1, 300)
        self.r1_spinbox.setValue(5)
        self.r1_spinbox.valueChanged.connect(self.update_durations)
        r1_layout.addWidget(self.r1_spinbox)
        duration_layout.addLayout(r1_layout)

        # Replicate 2 duration
        r2_layout = QHBoxLayout()
        r2_layout.addWidget(QLabel("Replicate 2:"))
        self.r2_spinbox = QSpinBox()
        self.r2_spinbox.setRange(1, 300)
        self.r2_spinbox.setValue(5)
        self.r2_spinbox.valueChanged.connect(self.update_durations)
        r2_layout.addWidget(self.r2_spinbox)
        duration_layout.addLayout(r2_layout)

        # Replicate 3 duration
        r3_layout = QHBoxLayout()
        r3_layout.addWidget(QLabel("Replicate 3:"))
        self.r3_spinbox = QSpinBox()
        self.r3_spinbox.setRange(1, 300)
        self.r3_spinbox.setValue(5)
        self.r3_spinbox.valueChanged.connect(self.update_durations)
        r3_layout.addWidget(self.r3_spinbox)
        duration_layout.addLayout(r3_layout)

        layout.addLayout(duration_layout)
        # Initialize the text output display
        self.text_area = QPlainTextEdit()
        self.text_area.setReadOnly(True)
        layout.addWidget(self.text_area)
        # Initialize the automation on/off toggle
        self.toggle_button = QPushButton("Start Automation")
        self.toggle_button.clicked.connect(self.toggle_automation)
        self.automation_on = False
        layout.addWidget(self.toggle_button)
        self.setLayout(layout)
        # Set size and bottom-right position
        self.window_width = 425
        self.window_height = 400
        screen_geometry: QRect = QApplication.primaryScreen().availableGeometry()
        x = screen_geometry.x() + screen_geometry.width() - self.window_width
        y = screen_geometry.y() + screen_geometry.height() - self.window_height
        self.setGeometry(x, y, self.window_width, self.window_height)
        # Redirect stdout
        sys.stdout = OutputRedirect(self.append_text)

    def lock_sample_inputs(self):
        """Lock all sample configuration inputs"""
        self.sample_name_input.setEnabled(False)
        self.last_sample_checkbox.setEnabled(False)
        self.repeat_count_spinbox.setEnabled(False)
        self.r1_spinbox.setEnabled(False)
        self.r2_spinbox.setEnabled(False)
        self.r3_spinbox.setEnabled(False)
        self.submit_button.setEnabled(False)

    def unlock_sample_inputs(self):
        """Unlock sample configuration inputs for next sample entry"""
        self.sample_name_input.setEnabled(True)
        self.last_sample_checkbox.setEnabled(True)
        self.repeat_count_spinbox.setEnabled(True)
        self.r1_spinbox.setEnabled(True)
        self.r2_spinbox.setEnabled(True)
        self.r3_spinbox.setEnabled(True)
        self.submit_button.setEnabled(True)

    def submit_sample(self):
        """Trigger sample submission from the GUI button"""
        self.update_sample_config()  # Capture latest field values before submitting
        threading.Thread(
            target=self.pump_controller.submit_sample, daemon=True
        ).start()

    def run_wash(self):
        """Trigger a 60-second wash cycle in a background thread"""
        self.pump_controller.log("Wash triggered - running pump for 60 seconds")
        threading.Thread(
            target=self.pump_controller.pump_sample, daemon=True
        ).start()

    def update_sample_config(self):
        """Update the pump controller with new sample configuration"""
        self.pump_controller.base_sample_name = self.sample_name_input.text().strip()
        self.pump_controller.repeat_count = self.repeat_count_spinbox.value()
        self.pump_controller.last_sample = self.last_sample_checkbox.isChecked()

    def update_durations(self):
        """Update the pump controller with new replicate durations"""
        self.pump_controller.replicate_durations = [
            self.r1_spinbox.value(),
            self.r2_spinbox.value(),
            self.r3_spinbox.value(),
        ]

    def append_text(self, text):
        self.text_area.appendPlainText(text.strip())

    def toggle_automation(self):
        self.automation_on = not self.automation_on
        if self.automation_on:
            self.lock_sample_inputs()
            self.toggle_button.setText("Stop Autopump")
            self.pump_controller.start_automation()
        else:
            self.lock_sample_inputs()
            self.toggle_button.setText("Start Autopump")
            self.pump_controller.stop_automation()


class TrayApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.log_window = LogWindow()
        # Setup tray icon
        self.tray_icon = QSystemTrayIcon()
        self.tray_icon.setIcon(QIcon.fromTheme("applications-system"))
        self.tray_icon.setToolTip("Pump Automation")
        # Right-click menu
        menu = QMenu()
        toggle_action = QAction("Show/Hide Log Window", self.tray_icon)
        toggle_action.triggered.connect(self.toggle_window)
        menu.addAction(toggle_action)
        quit_action = QAction("Exit", self.tray_icon)
        quit_action.triggered.connect(self.quit)
        menu.addAction(quit_action)
        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self.icon_clicked)
        self.tray_icon.show()
        # Test output
        print("Tray icon initialized.")
        QTimer.singleShot(1000, lambda: print("Pump automation system ready."))

    def toggle_window(self):
        if self.log_window.isVisible():
            self.log_window.hide()
        else:
            self.log_window.show()

    def icon_clicked(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            self.toggle_window()

    def quit(self):
        if self.log_window.pump_controller.automation_active:
            self.log_window.pump_controller.stop_automation()
            self.tray_icon.hide()
            self.app.quit()

    def run(self):
        sys.exit(self.app.exec())


if __name__ == "__main__":
    time.sleep(2)
    TrayApp().run()