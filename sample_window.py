"""
Test script to simulate the Start Sample Window using native Win32 controls
This window has NO Continue button, only an Accept button
Compatible with pywinauto for automation
"""
import win32gui
import win32con
import win32api
import sys
import time


class StartSampleWindow:
    """Native Windows dialog with two Edit controls and an Accept button"""

    def __init__(self):
        self.hwnd = None
        self.edit1_hwnd = None  # Sample name edit box
        self.edit2_hwnd = None  # Last sample (Yes/No) edit box
        self.button_hwnd = None
        self.running = False
        self.sample_name = ""
        self.last_sample = ""

        # Window class name
        self.class_name = "StartSampleWindow"

    def create_window(self):
        """Create the native Windows dialog"""
        # Register window class
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = self._wnd_proc
        wc.lpszClassName = self.class_name
        wc.hInstance = win32api.GetModuleHandle(None)
        wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        wc.hbrBackground = win32con.COLOR_WINDOW + 1

        try:
            win32gui.RegisterClass(wc)
        except Exception as e:
            # Class might already be registered
            pass

        # Create main window
        self.hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_DLGMODALFRAME,  # Extended style for dialog look
            self.class_name,
            "Operator Request",  # Window title - matches what autopump looks for
            win32con.WS_OVERLAPPED | win32con.WS_CAPTION | win32con.WS_SYSMENU | win32con.WS_VISIBLE,
            300, 300,  # x, y position
            400, 250,  # width, height
            0, 0,
            wc.hInstance,
            None
        )

        # Create label 1
        label1_hwnd = win32gui.CreateWindowEx(
            0,
            "STATIC",
            "Enter Sample Name:",
            win32con.WS_CHILD | win32con.WS_VISIBLE,
            20, 20, 350, 20,
            self.hwnd, 0, wc.hInstance, None
        )

        # Create Edit control 1 - Sample Name
        self.edit1_hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_CLIENTEDGE,
            "EDIT",  # Native Windows EDIT class
            "",
            win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_BORDER | win32con.ES_AUTOHSCROLL,
            20, 45, 350, 25,
            self.hwnd, 1001, wc.hInstance, None
        )

        # Create label 2
        label2_hwnd = win32gui.CreateWindowEx(
            0,
            "STATIC",
            "Is this the last sample? (Yes/No):",
            win32con.WS_CHILD | win32con.WS_VISIBLE,
            20, 85, 350, 20,
            self.hwnd, 0, wc.hInstance, None
        )

        # Create Edit control 2 - Last Sample (Yes/No)
        self.edit2_hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_CLIENTEDGE,
            "EDIT",  # Native Windows EDIT class
            "",
            win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_BORDER | win32con.ES_AUTOHSCROLL,
            20, 110, 350, 25,
            self.hwnd, 1002, wc.hInstance, None
        )

        # Create Accept button
        self.button_hwnd = win32gui.CreateWindowEx(
            0,
            "BUTTON",
            "Accept",
            win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_PUSHBUTTON,
            150, 160, 100, 30,
            self.hwnd, 1003, wc.hInstance, None
        )

        # Set default font for better appearance
        # DEFAULT_GUI_FONT = 17 (constant value)
        font = win32gui.GetStockObject(17)
        win32gui.SendMessage(label1_hwnd, win32con.WM_SETFONT, font, True)
        win32gui.SendMessage(self.edit1_hwnd, win32con.WM_SETFONT, font, True)
        win32gui.SendMessage(label2_hwnd, win32con.WM_SETFONT, font, True)
        win32gui.SendMessage(self.edit2_hwnd, win32con.WM_SETFONT, font, True)
        win32gui.SendMessage(self.button_hwnd, win32con.WM_SETFONT, font, True)

        # Show window
        win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)
        win32gui.UpdateWindow(self.hwnd)

        return self.hwnd

    def _wnd_proc(self, hwnd, msg, wparam, lparam):
        """Window procedure to handle messages"""
        if msg == win32con.WM_COMMAND:
            # Button click - extract low word from wparam
            control_id = wparam & 0xFFFF  # LOWORD equivalent
            if control_id == 1003:  # Accept button ID
                self._on_accept()
                return 0
        elif msg == win32con.WM_CLOSE:
            self._on_close()
            return 0
        elif msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0

        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def _on_accept(self):
        """Handle Accept button click"""
        # Get text from edit controls
        self.sample_name = win32gui.GetWindowText(self.edit1_hwnd)
        self.last_sample = win32gui.GetWindowText(self.edit2_hwnd)

        print(f"Accept clicked - Sample: {self.sample_name}, Last Sample: {self.last_sample}")

        # Close the window
        self._on_close()

    def _on_close(self):
        """Close the window"""
        self.running = False
        win32gui.DestroyWindow(self.hwnd)

    def run(self):
        """Run the message loop"""
        self.running = True
        self.create_window()

        # Message loop
        while self.running:
            try:
                win32gui.PumpWaitingMessages()
                time.sleep(0.01)
            except Exception as e:
                print(f"Message loop error: {e}")
                break


if __name__ == "__main__":
    import threading

    # Create and show the window in a separate thread
    window = StartSampleWindow()

    def run_window():
        window.run()

    window_thread = threading.Thread(target=run_window, daemon=True)
    window_thread.start()

    # Give window time to fully create
    time.sleep(1)

    # Test pywinauto detection
    try:
        from pywinauto import Desktop
        print("\n=== PyWinAuto Detection Test ===")
        windows = Desktop(backend="win32").windows()
        for win in windows:
            if "Operator Request" in win.window_text():
                print(f"✓ Found window: {win.window_text()}")
                print(f"✓ Window class: {win.class_name()}")
                print("\nChild controls:")
                for child in win.iter_children():
                    print(f"  - Class: '{child.friendly_class_name()}' | Text: '{child.window_text()}'")

                # Check specifically for Edit controls
                edit_controls = []
                for child in win.iter_children():
                    if child.friendly_class_name() == 'Edit':
                        edit_controls.append(child)

                print(f"\n{'='*50}")
                print(f"✓ Found {len(edit_controls)} Edit controls")
                if len(edit_controls) >= 2:
                    print("✓✓✓ SUCCESS: Both edit boxes detected as 'Edit' class!")
                    print("✓✓✓ This window is compatible with autopump_v2_0.py")
                else:
                    print("✗ WARNING: Not enough Edit controls detected")
                print(f"{'='*50}\n")
                break
    except Exception as e:
        print(f"Test error: {e}")

    # Keep main thread alive to see the window
    print("Window is running. Close the window to exit.")
    try:
        while window.running:
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nExiting...")
        window._on_close()