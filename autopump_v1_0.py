import os
import sys

# Set DPI awareness before Qt imports to avoid warnings
if sys.platform == "win32":
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"

from PyQt6.QtWidgets import (
    QApplication, QPlainTextEdit, QVBoxLayout, QWidget,
    QSystemTrayIcon, QMenu, QPushButton
)
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtCore import QTimer, Qt, QRect, pyqtSignal, QObject
from typing import Any
import sys
import io
import time
import serial
from pywinauto import Desktop
import keyboard
import threading
import pythoncom  # Add this for COM initialization


class OutputRedirect(io.StringIO):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def write(self, text):
        self.callback(text)


class PumpController(QObject):
    # Signal for thread-safe logging
    log_message = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.SERIAL_PORT = "COM5"
        self.runtime_sec = 5
        self.window_name = "Operator Request"
        self.automation_active = False
        self.f8_listener_thread = None
        self.window_monitor_thread = None

    def log(self, message):
        """Thread-safe logging method"""
        self.log_message.emit(message)  # type: ignore

    def pump_sample(self):
        """Execute a pump sample cycle via serial communication"""
        try:
            with serial.Serial(self.SERIAL_PORT, 1200, timeout=1) as ser:
                ser.write(f"START:{self.runtime_sec}\n".encode())
                self.log(f"Command sent: START:{self.runtime_sec}")
                while True:
                    line = ser.readline().decode().strip()
                    if line:
                        self.log(f"Pico says: {line}")
                    if line == "Pump stopped":
                        break
                    time.sleep(0.5)
        except Exception as e:
            self.log(f"Pump sample error: {e}")

    def click_continue_button(self, window):
        """Click the continue button in operator request dialogs"""
        try:
            for ctrl in window.descendants():
                if ctrl.friendly_class_name() == 'Button':
                    ctrl.click_input()
                    self.log("Clicked 'Continue'")
                    return
            self.log("Continue button not available")
        except Exception as exc:
            self.log(f"Failed to click Continue: {exc}")

    def is_correct_operator_request(self, window):
        """Check if the window contains the correct operator request text"""
        try:
            for ctrl in window.descendants():
                if ctrl.friendly_class_name() == 'Edit':
                    text = ctrl.window_text().strip()
                    if "Pull sample through flow cell" in text:
                        return True
            return False
        except Exception as exc:
            self.log(f"Window check error: {exc}")
            return False

    def listen_for_f8(self):
        """Background thread to listen for F8 key presses"""
        cooldown_sec = self.runtime_sec + 1
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
                    self.log(f"F8 pressed too soon. Wait {cooldown_sec - (now - last_trigger_time):.1f} more seconds.")
            except Exception as e:
                self.log(f"F8 listener error: {e}")
                time.sleep(1)

    def monitor_windows(self):
        """Background thread to monitor for operator request windows"""
        # Initialize COM for this thread
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
                            if self.is_correct_operator_request(win):
                                self.log("Correct 'Operator Request' detected!")
                                self.pump_sample()
                                self.click_continue_button(win)
                                time.sleep(3)
                                break
                    time.sleep(1)
                except Exception as e:
                    self.log(f"Window monitor error: {e}")
                    time.sleep(1)
        finally:
            # Clean up COM
            pythoncom.CoUninitialize()

    def start_automation(self):
        """Start the pump automation system"""
        if self.automation_active:
            self.log("Automation already running")
            return

        self.automation_active = True
        self.log("Starting pump automation...")

        # Start F8 listener thread
        self.f8_listener_thread = threading.Thread(target=self.listen_for_f8, daemon=True)
        self.f8_listener_thread.start()

        # Start window monitor thread
        self.window_monitor_thread = threading.Thread(target=self.monitor_windows, daemon=True)
        self.window_monitor_thread.start()

        self.log("Automation threads started. F8 manual trigger and window monitoring active.")

    def stop_automation(self):
        """Stop the pump automation system"""
        if not self.automation_active:
            self.log("Automation already stopped")
            return

        self.log("Stopping pump automation...")
        self.automation_active = False

        # Threads will stop on their own due to the automation_active flag
        self.log("Automation stopped. All monitoring threads will terminate.")


class LogWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowFlags(Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowTitle("Automation Log")

        # Initialize pump controller
        self.pump_controller = PumpController()

        # Connect pump controller's log signal to our GUI update method
        self.pump_controller.log_message.connect(self.append_text)  # type: ignore

        # initialize the text output display
        self.text_area = QPlainTextEdit()
        self.text_area.setReadOnly(True)

        # initialize the automation on/off toggle
        self.toggle_button = QPushButton("Start Automation")
        self.toggle_button.clicked.connect(self.toggle_automation)  # type: ignore
        self.automation_on = False

        layout = QVBoxLayout()
        layout.addWidget(self.text_area)
        layout.addWidget(self.toggle_button)
        self.setLayout(layout)

        # Set size and bottom-right position
        self.window_width = 400
        self.window_height = 200
        screen_geometry: QRect = QApplication.primaryScreen().availableGeometry()
        x = screen_geometry.x() + screen_geometry.width() - self.window_width
        y = screen_geometry.y() + screen_geometry.height() - self.window_height
        self.setGeometry(x, y, self.window_width, self.window_height)

        # Redirect stdout
        sys.stdout = OutputRedirect(self.append_text)

    def append_text(self, text):
        self.text_area.appendPlainText(text.strip())

    def toggle_automation(self):
        self.automation_on = not self.automation_on
        if self.automation_on:
            self.toggle_button.setText("Stop Autopump")
            self.pump_controller.start_automation()
        else:
            self.toggle_button.setText("Start Autopump")
            self.pump_controller.stop_automation()


class TrayApp:
    def __init__(self):
        self.app = QApplication(sys.argv)
        self.log_window = LogWindow()

        # Setup tray icon
        self.tray_icon = QSystemTrayIcon()
        self.tray_icon.setIcon(QIcon.fromTheme("applications-system"))  # fallback icon
        self.tray_icon.setToolTip("Pump Automation")

        # Right-click menu
        menu = QMenu()
        toggle_action = QAction("Show/Hide Log Window", self.tray_icon)
        toggle_action.triggered.connect(self.toggle_window)  # type: ignore
        menu.addAction(toggle_action)

        quit_action = QAction("Exit", self.tray_icon)
        quit_action.triggered.connect(self.quit)  # type: ignore
        menu.addAction(quit_action)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self.icon_clicked)  # type: ignore

        self.tray_icon.show()

        # Test output - use thread-safe logging
        print("Tray icon initialized.")
        QTimer.singleShot(1000, lambda: print("Pump automation system ready."))

    def toggle_window(self):
        if self.log_window.isVisible():
            self.log_window.hide()
        else:
            self.log_window.show()

    def icon_clicked(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:  # Left click
            self.toggle_window()

    def quit(self):
        # Stop automation before quitting
        if self.log_window.pump_controller.automation_active:
            self.log_window.pump_controller.stop_automation()
        self.tray_icon.hide()
        self.app.quit()

    def run(self):
        sys.exit(self.app.exec())


if __name__ == "__main__":
    # Add initialization delay like the original
    time.sleep(2)
    TrayApp().run()