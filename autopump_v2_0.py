import os
import sys

# Set DPI awareness before Qt imports to avoid warnings
if sys.platform == "win32":
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "0"

from PyQt6.QtWidgets import (
    QApplication, QPlainTextEdit, QVBoxLayout, QWidget,
    QSystemTrayIcon, QMenu, QPushButton, QTableWidget,
    QTableWidgetItem, QHBoxLayout, QHeaderView
)
from PyQt6.QtGui import QIcon, QAction, QColor
from PyQt6.QtCore import QTimer, Qt, QRect, pyqtSignal, QObject, QDateTime
from typing import Any
import io
import time
import serial
from pywinauto import Desktop
import keyboard
import threading
import pythoncom  # Add this for COM initialization
import pywinauto


class OutputRedirect(io.StringIO):
    def __init__(self, callback):
        super().__init__()
        self.callback = callback

    def write(self, text):
        self.callback(text)


class SampleQueue(QObject):
    # Signals for communication with other components
    sample_started = pyqtSignal(dict)  # Emitted when sample starts
    sample_completed = pyqtSignal(dict)  # Emitted when sample completes
    queue_updated = pyqtSignal()  # Emitted when queue changes

    def __init__(self):
        super().__init__()
        self.samples = []  # List of sample dictionaries
        self.current_sample_index = None
        self.sample_types = ["Standard", "Blank", "QC", "Unknown"]

        # Column definitions - added cycle durations
        self.columns = [
            "Queue #",
            "Sample ID",
            "Sample Type",
            "Rep 1 (sec)",
            "Rep 2 (sec)",
            "Rep 3 (sec)",
            "Status",
            "Start Time",
            "Notes"
        ]

    def add_sample(self, sample_id="", sample_type="Standard", rep1_duration=60, rep2_duration=60, rep3_duration=60,
                   notes=""):
        """Add a new sample to the queue"""
        sample = {
            "queue_num": len(self.samples) + 1,
            "sample_id": sample_id or f"Sample_{len(self.samples) + 1:03d}",
            "sample_type": sample_type,
            "rep1_duration": rep1_duration,
            "rep2_duration": rep2_duration,
            "rep3_duration": rep3_duration,
            "status": "Pending",
            "start_time": "",
            "notes": notes,
            "created_time": QDateTime.currentDateTime().toString(),
            "current_replicate": 0  # Track which replicate we're on (0 = not started, 1-3 = replicates)
        }
        self.samples.append(sample)
        self._renumber_queue()
        self.queue_updated.emit()  # type: ignore
        return len(self.samples) - 1  # Return index of added sample

    def remove_sample(self, index):
        """Remove sample at given index"""
        if 0 <= index < len(self.samples):
            removed = self.samples.pop(index)
            self._renumber_queue()
            self.queue_updated.emit()  # type: ignore
            return removed
        return None

    def move_sample(self, from_index, to_index):
        """Move sample from one position to another"""
        if 0 <= from_index < len(self.samples) and 0 <= to_index < len(self.samples):
            sample = self.samples.pop(from_index)
            self.samples.insert(to_index, sample)
            self._renumber_queue()
            self.queue_updated.emit()  # type: ignore

    def get_next_sample(self):
        """Get the next pending sample to run"""
        for i, sample in enumerate(self.samples):
            if sample["status"] == "Pending":
                return i, sample
        return None, None

    def start_sample(self, index):
        """Mark sample as running and set start time"""
        if 0 <= index < len(self.samples):
            self.samples[index]["status"] = "Running"
            self.samples[index]["start_time"] = QDateTime.currentDateTime().toString("hh:mm:ss")
            self.samples[index]["current_replicate"] = 1  # Starting replicate 1
            self.current_sample_index = index
            self.sample_started.emit(self.samples[index])  # type: ignore
            self.queue_updated.emit()  # type: ignore

    def advance_replicate(self, index):
        """Move to the next replicate for a sample"""
        if 0 <= index < len(self.samples):
            current_rep = self.samples[index]["current_replicate"]
            if current_rep < 3:
                self.samples[index]["current_replicate"] = current_rep + 1
                self.queue_updated.emit()  # type: ignore
                return self.samples[index]["current_replicate"]
            return None  # All replicates done

    def get_current_replicate_duration(self, index):
        """Get the duration for the current replicate"""
        if 0 <= index < len(self.samples):
            current_rep = self.samples[index]["current_replicate"]
            if current_rep == 1:
                return self.samples[index]["rep1_duration"]
            elif current_rep == 2:
                return self.samples[index]["rep2_duration"]
            elif current_rep == 3:
                return self.samples[index]["rep3_duration"]
        return 30  # Default

    def complete_sample(self, index, success=True):
        """Mark sample as complete or error"""
        if 0 <= index < len(self.samples):
            self.samples[index]["status"] = "Complete" if success else "Error"
            if self.current_sample_index == index:
                self.current_sample_index = None
            self.sample_completed.emit(self.samples[index])  # type: ignore
            self.queue_updated.emit()  # type: ignore

    def update_sample(self, index, field, value):
        """Update a specific field of a sample"""
        if 0 <= index < len(self.samples) and field in self.samples[index]:
            self.samples[index][field] = value
            self.queue_updated.emit()  # type: ignore

    def get_sample_count_by_status(self, status):
        """Get count of samples with given status"""
        return sum(1 for sample in self.samples if sample["status"] == status)

    def clear_completed_samples(self):
        """Remove all completed samples from queue"""
        self.samples = [s for s in self.samples if s["status"] not in ["Complete"]]
        self._renumber_queue()
        self.queue_updated.emit()  # type: ignore

    def get_sample_info(self, index):
        """Get sample information by index"""
        if 0 <= index < len(self.samples):
            return self.samples[index].copy()
        return None

    def get_all_samples(self):
        """Get all samples (copy of the list)"""
        return [sample.copy() for sample in self.samples]

    def _renumber_queue(self):
        """Renumber all samples in queue order"""
        for i, sample in enumerate(self.samples):
            sample["queue_num"] = i + 1


class PumpController(QObject):
    # Signal for thread-safe logging
    log_message = pyqtSignal(str)

    def __init__(self, sample_queue=None):
        super().__init__()
        self.SERIAL_PORT = "COM5"
        self.runtime_sec = 60  # Default, will be overridden by sample durations
        self.window_name = "Operator Request"
        self.automation_active = False
        self.f8_listener_thread = None
        self.window_monitor_thread = None
        self.sample_queue = sample_queue
        self.current_sample_index = None

    def log(self, message):
        """Thread-safe logging method"""
        self.log_message.emit(message)  # type: ignore

    def pump_sample(self, duration_sec=None):
        """Execute a pump sample cycle via serial communication"""
        if duration_sec is None:
            duration_sec = self.runtime_sec

        try:
            with serial.Serial(self.SERIAL_PORT, 1200, timeout=1) as ser:
                ser.write(f"START:{duration_sec}\n".encode())
                self.log(f"Command sent: START:{duration_sec}")
                while True:
                    line = ser.readline().decode().strip()
                    if line:
                        self.log(f"Pico says: {line}")
                    if line == "Pump stopped":
                        break
                    time.sleep(0.5)
        except Exception as e:
            self.log(f"Pump sample error: {e}")

    def is_start_sample_window(self, window):
        """Check if this is the start sample window (no Continue button)"""
        try:
            has_continue = False
            for ctrl in window.descendants():
                if ctrl.friendly_class_name() == 'Button':
                    button_text = ctrl.window_text().lower()
                    if 'continue' in button_text:
                        has_continue = True
                        break

            # This is a start sample window if it has NO continue button
            # but IS an Operator Request window
            return not has_continue and self.window_name in window.window_text()
        except Exception as exc:
            self.log(f"Window check error: {exc}")
            return False

    def fill_start_sample_window(self, window, sample_data, is_last_sample):
        """Fill in the sample name and last sample fields"""
        try:
            edit_controls = []
            for child in window.iter_children():
                if child.friendly_class_name() == 'Edit':
                    edit_controls.append(child)

            if len(edit_controls) >= 2:
                # First text box: Sample name
                edit_controls[0].set_focus()
                edit_controls[0].set_text(sample_data['sample_id'])
                self.log(f"Filled sample name: {sample_data['sample_id']}")

                # Second text box: Yes/No for last sample
                last_sample_text = "Yes" if is_last_sample else "No"
                edit_controls[1].set_focus()
                edit_controls[1].set_text(last_sample_text)
                self.log(f"Filled last sample: {last_sample_text}")

                return True
            else:
                self.log(f"Could not find both text entry boxes (found {len(edit_controls)})")
                return False

        except Exception as exc:
            self.log(f"Error filling start sample window: {exc}")
            return False

    def click_continue_button(self, window):
        """Click the continue button in operator request dialogs"""
        try:
            for ctrl in window.descendants():
                if ctrl.friendly_class_name() == 'Button':
                    button_text = ctrl.window_text().lower()
                    if 'continue' or 'accept' in button_text:
                        ctrl.set_focus()
                        ctrl.click_input()
                        self.log("Clicked 'Continue' or 'Accept'")
                        return
            self.log("Continue or accept button not available")
        except Exception as exc:
            self.log(f"Failed to click Continue or Accept: {exc}")

    def is_correct_operator_request(self, window):
        """Check if the window contains the correct operator request text (replicate pump window)"""
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
                    self.pump_sample()  # Manual pump with default duration
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

                            # Check if this is a start sample window (no Continue button)
                            if self.is_start_sample_window(win):
                                self.log("Start sample window detected!")

                                # Get next sample from queue
                                if self.sample_queue:
                                    next_idx, next_sample = self.sample_queue.get_next_sample()
                                    if next_sample:
                                        # Check if this is the last sample
                                        is_last = (next_idx == len(self.sample_queue.samples) - 1)

                                        # Mark sample as running
                                        self.sample_queue.start_sample(next_idx)
                                        self.current_sample_index = next_idx

                                        # Fill in sample info and click Accept
                                        if self.fill_start_sample_window(win, next_sample, is_last):
                                            self.log(f"Started sample: {next_sample['sample_id']}")
                                            self.click_continue_button(win)
                                        time.sleep(2)
                                        break
                                    else:
                                        self.log("No pending samples in queue")

                            # Check if this is a replicate pump window (has Continue button)
                            elif self.is_correct_operator_request(win):
                                self.log("Replicate pump window detected!")

                                # Get current sample and replicate info
                                if self.sample_queue and self.current_sample_index is not None:
                                    # Check if the sample still exists in the queue
                                    if self.current_sample_index < len(self.sample_queue.samples):
                                        current_sample = self.sample_queue.samples[self.current_sample_index]
                                        rep_num = current_sample["current_replicate"]
                                        duration = self.sample_queue.get_current_replicate_duration(
                                            self.current_sample_index)

                                        self.log(f"Running replicate {rep_num} for {duration} seconds")

                                        # Run pump with this replicate's duration
                                        self.pump_sample(duration)

                                        # Click continue
                                        self.click_continue_button(win)

                                        # Advance to next replicate or complete sample
                                        next_rep = self.sample_queue.advance_replicate(self.current_sample_index)
                                        if next_rep is None:
                                            # All replicates done, mark as complete
                                            self.sample_queue.complete_sample(self.current_sample_index, True)
                                            self.log(f"Sample {current_sample['sample_id']} completed")
                                            self.current_sample_index = None
                                        else:
                                            self.log(f"Ready for replicate {next_rep}")

                                        time.sleep(3)
                                        break
                                    else:
                                        # Sample was removed from queue, continue with manual pump
                                        self.log("Sample was removed from queue - continuing with default duration")
                                        self.current_sample_index = None
                                        self.pump_sample()
                                        self.click_continue_button(win)
                                        time.sleep(3)
                                        break
                                else:
                                    # Manual pump if no queue
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

        # Initialize sample queue (Step 1: Data only, no UI)
        self.sample_queue = SampleQueue()

        # Initialize pump controller with sample queue reference
        self.pump_controller = PumpController(self.sample_queue)

        # Connect pump controller's log signal to our GUI update method
        self.pump_controller.log_message.connect(self.append_text)  # type: ignore

        # Create main horizontal layout (log on left, table on right)
        main_layout = QHBoxLayout()

        # Left side: Log output and button
        left_layout = QVBoxLayout()

        # initialize the text output display
        self.text_area = QPlainTextEdit()
        self.text_area.setReadOnly(True)

        # initialize the automation on/off toggle
        self.toggle_button = QPushButton("Start Automation")
        self.toggle_button.clicked.connect(self.toggle_automation)  # type: ignore
        self.automation_on = False

        left_layout.addWidget(self.text_area)
        left_layout.addWidget(self.toggle_button)

        # Right side: Sample queue table
        right_layout = QVBoxLayout()

        # Add control buttons above the table
        button_layout = QHBoxLayout()
        self.add_sample_btn = QPushButton("Add Sample")
        self.remove_sample_btn = QPushButton("Remove Selected")
        self.clear_completed_btn = QPushButton("Clear Completed")

        self.add_sample_btn.clicked.connect(self.add_sample_to_queue)  # type: ignore
        self.remove_sample_btn.clicked.connect(self.remove_selected_sample)  # type: ignore
        self.clear_completed_btn.clicked.connect(self.clear_completed_samples)  # type: ignore

        button_layout.addWidget(self.add_sample_btn)
        button_layout.addWidget(self.remove_sample_btn)
        button_layout.addWidget(self.clear_completed_btn)
        button_layout.addStretch()

        # Create the table widget
        self.table = QTableWidget()
        self.table.setColumnCount(len(self.sample_queue.columns))
        self.table.setHorizontalHeaderLabels(self.sample_queue.columns)
        self.table.setRowCount(10)  # 10 empty rows for user input

        # Configure table appearance
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        # Connect cell changes to update queue
        self.table.cellChanged.connect(self.on_table_cell_changed)  # type: ignore

        # Connect sample queue signals to refresh table
        self.sample_queue.queue_updated.connect(self.refresh_table_from_queue)  # type: ignore

        # Initialize empty rows with Queue # pre-filled
        self.is_refreshing = False  # Flag to prevent recursive updates
        for row in range(10):
            # Queue # column - auto-numbered and read-only
            queue_num_item = QTableWidgetItem(str(row + 1))
            queue_num_item.setFlags(queue_num_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, queue_num_item)

            # Initialize other columns with empty cells
            for col in range(1, len(self.sample_queue.columns)):
                self.table.setItem(row, col, QTableWidgetItem(""))

        right_layout.addLayout(button_layout)
        right_layout.addWidget(self.table)

        # Add both sides to main layout
        main_layout.addLayout(left_layout, stretch=1)
        main_layout.addLayout(right_layout, stretch=1)

        self.setLayout(main_layout)

        # Set size and bottom-right position (wider to accommodate both panels)
        self.window_width = 900  # Wider for side-by-side layout
        self.window_height = 400  # Taller for table
        screen_geometry: QRect = QApplication.primaryScreen().availableGeometry()
        x = screen_geometry.x() + screen_geometry.width() - self.window_width
        y = screen_geometry.y() + screen_geometry.height() - self.window_height
        self.setGeometry(x, y, self.window_width, self.window_height)

        # Redirect stdout
        sys.stdout = OutputRedirect(self.append_text)

        # Test the sample queue and display in table
        self.test_sample_queue()

        # Initial table refresh to show test data
        self.refresh_table_from_queue()

    def refresh_table_from_queue(self):
        """Refresh the table display from SampleQueue data"""
        self.is_refreshing = True  # Prevent recursive updates

        # Clear existing table
        self.table.setRowCount(0)

        # Add rows for each sample in queue
        for row, sample in enumerate(self.sample_queue.samples):
            self.table.insertRow(row)

            # Queue #
            queue_item = QTableWidgetItem(str(sample["queue_num"]))
            queue_item.setFlags(queue_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 0, queue_item)

            # Sample ID
            self.table.setItem(row, 1, QTableWidgetItem(sample["sample_id"]))

            # Sample Type
            self.table.setItem(row, 2, QTableWidgetItem(sample["sample_type"]))

            # Rep 1 Duration
            self.table.setItem(row, 3, QTableWidgetItem(str(sample["rep1_duration"])))

            # Rep 2 Duration
            self.table.setItem(row, 4, QTableWidgetItem(str(sample["rep2_duration"])))

            # Rep 3 Duration
            self.table.setItem(row, 5, QTableWidgetItem(str(sample["rep3_duration"])))

            # Status - make read-only and color-code
            status_text = sample["status"]
            if sample["status"] == "Running":
                rep_num = sample.get("current_replicate", 0)
                if rep_num > 0:
                    status_text = f"Running (Rep {rep_num}/3)"

            status_item = QTableWidgetItem(status_text)
            status_item.setFlags(status_item.flags() & ~Qt.ItemFlag.ItemIsEditable)

            # Color code status
            if sample["status"] == "Running":
                status_item.setBackground(QColor(255, 255, 0))  # Yellow
            elif sample["status"] == "Complete":
                status_item.setBackground(QColor(144, 238, 144))  # Light green
            elif sample["status"] == "Error":
                status_item.setBackground(QColor(255, 182, 193))  # Light red

            self.table.setItem(row, 6, status_item)

            # Start Time - read-only
            time_item = QTableWidgetItem(sample["start_time"])
            time_item.setFlags(time_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            self.table.setItem(row, 7, time_item)

            # Notes
            self.table.setItem(row, 8, QTableWidgetItem(sample["notes"]))

        # Add empty rows to reach minimum of 10 rows
        current_rows = self.table.rowCount()
        if current_rows < 10:
            for row in range(current_rows, 10):
                self.table.insertRow(row)
                # Queue # for empty rows
                queue_item = QTableWidgetItem(str(row + 1))
                queue_item.setFlags(queue_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                self.table.setItem(row, 0, queue_item)

                # Empty cells for other columns
                for col in range(1, len(self.sample_queue.columns)):
                    self.table.setItem(row, col, QTableWidgetItem(""))

        self.is_refreshing = False

    def on_table_cell_changed(self, row, col):
        """Handle user editing cells in the table"""
        if self.is_refreshing:
            return  # Don't process changes during refresh

        # Editable columns: Sample ID (1), Sample Type (2), Rep durations (3,4,5), Notes (8)
        if col in [1, 2, 3, 4, 5, 8]:
            item = self.table.item(row, col)
            if item is None:
                return

            new_value = item.text().strip()

            # Check if this row has a sample in the queue
            if row < len(self.sample_queue.samples):
                # Update existing sample
                field_map = {
                    1: "sample_id",
                    2: "sample_type",
                    3: "rep1_duration",
                    4: "rep2_duration",
                    5: "rep3_duration",
                    8: "notes"
                }
                field_name = field_map[col]

                # Convert duration values to integers
                if col in [3, 4, 5]:
                    try:
                        new_value = int(new_value) if new_value else 60
                    except ValueError:
                        new_value = 60
                        self.table.item(row, col).setText("60")

                self.sample_queue.update_sample(row, field_name, new_value)
                print(f"Updated sample {row + 1}: {field_name} = {new_value}")
            else:
                # This is an empty row - check if user entered data
                sample_id = self.table.item(row, 1).text().strip() if self.table.item(row, 1) else ""
                sample_type = self.table.item(row, 2).text().strip() if self.table.item(row, 2) else ""
                rep1 = self.table.item(row, 3).text().strip() if self.table.item(row, 3) else "60"
                rep2 = self.table.item(row, 4).text().strip() if self.table.item(row, 4) else "60"
                rep3 = self.table.item(row, 5).text().strip() if self.table.item(row, 5) else "60"
                notes = self.table.item(row, 8).text().strip() if self.table.item(row, 8) else ""

                # Only create new sample if Sample ID is not empty
                if sample_id:
                    if not sample_type:
                        sample_type = "Standard"  # Default type

                    # Convert durations to integers
                    try:
                        rep1_dur = int(rep1) if rep1 else 60
                        rep2_dur = int(rep2) if rep2 else 60
                        rep3_dur = int(rep3) if rep3 else 60
                    except ValueError:
                        rep1_dur = rep2_dur = rep3_dur = 60

                    self.sample_queue.add_sample(sample_id, sample_type, rep1_dur, rep2_dur, rep3_dur, notes)
                    print(f"Added new sample: {sample_id}")

    def add_sample_to_queue(self):
        """Add a new empty sample to the queue"""
        new_id = f"Sample_{len(self.sample_queue.samples) + 1:03d}"
        self.sample_queue.add_sample(new_id, "Standard", "")
        print(f"Added sample: {new_id}")

    def remove_selected_sample(self):
        """Remove the currently selected sample from queue"""
        current_row = self.table.currentRow()
        if current_row >= 0 and current_row < len(self.sample_queue.samples):
            removed = self.sample_queue.remove_sample(current_row)
            if removed:
                print(f"Removed sample: {removed['sample_id']}")
        else:
            print("No sample selected or row is empty")

    def clear_completed_samples(self):
        """Clear all completed samples from the queue"""
        completed_count = self.sample_queue.get_sample_count_by_status("Complete")
        self.sample_queue.clear_completed_samples()
        print(f"Cleared {completed_count} completed samples")

    def test_sample_queue(self):
        """Test the sample queue functionality"""
        # Add some test samples with replicate durations
        idx1 = self.sample_queue.add_sample("STD-001", "Standard", 60, 60, 60, "Calibration standard")
        idx2 = self.sample_queue.add_sample("BLK-001", "Blank", 45, 45, 45, "Method blank")
        idx3 = self.sample_queue.add_sample("QC-001", "QC", 30, 30, 30, "Quality control check")

        # Test logging queue status
        print(f"Sample queue initialized with {len(self.sample_queue.samples)} samples")

        # Test getting next sample
        next_idx, next_sample = self.sample_queue.get_next_sample()
        if next_sample:
            print(f"Next sample to run: {next_sample['sample_id']} ({next_sample['sample_type']})")

        # Test status counts
        pending_count = self.sample_queue.get_sample_count_by_status("Pending")
        print(f"Pending samples: {pending_count}")

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