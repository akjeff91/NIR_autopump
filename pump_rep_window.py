"""
Test script to simulate the Replicate Pump Window using native Win32 controls
This window HAS a Continue button and the "Pull sample through flow cell" text
Compatible with pywinauto for automation
"""
import win32gui
import win32con
import win32api
import sys
import time


class ReplicatePumpWindow:
    """Native Windows dialog with Edit control and Continue button"""

    def __init__(self):
        self.hwnd = None
        self.edit_hwnd = None  # Edit control with instruction text
        self.button_hwnd = None  # Continue button
        self.running = False

        # Window class name
        self.class_name = "ReplicatePumpWindow"

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
            win32con.WS_EX_DLGMODALFRAME,
            self.class_name,
            "Operator Request",  # Window title - matches what autopump looks for
            win32con.WS_OVERLAPPED | win32con.WS_CAPTION | win32con.WS_SYSMENU | win32con.WS_VISIBLE,
            300, 300,  # x, y position
            400, 250,  # width, height
            0, 0,
            wc.hInstance,
            None
        )

        # Create instruction label
        label1_hwnd = win32gui.CreateWindowEx(
            0,
            "STATIC",
            "Operator Action Required:",
            win32con.WS_CHILD | win32con.WS_VISIBLE,
            20, 20, 350, 20,
            self.hwnd, 0, wc.hInstance, None
        )

        # Create Edit control with instruction text
        # This is what autopump looks for: "Pull sample through flow cell"
        self.edit_hwnd = win32gui.CreateWindowEx(
            win32con.WS_EX_CLIENTEDGE,
            "EDIT",  # Native Windows EDIT class
            "Pull sample through flow cell",  # The text autopump searches for
            win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.WS_BORDER |
            win32con.ES_MULTILINE | win32con.ES_READONLY | win32con.ES_AUTOHSCROLL,
            20, 50, 350, 60,
            self.hwnd, 2001, wc.hInstance, None
        )

        # Create status label
        label2_hwnd = win32gui.CreateWindowEx(
            0,
            "STATIC",
            "Waiting for pump activation...",
            win32con.WS_CHILD | win32con.WS_VISIBLE,
            20, 125, 350, 20,
            self.hwnd, 0, wc.hInstance, None
        )

        # Create Continue button (this is what differentiates it from start sample window)
        self.button_hwnd = win32gui.CreateWindowEx(
            0,
            "BUTTON",
            "Continue",  # Button text that autopump looks for
            win32con.WS_CHILD | win32con.WS_VISIBLE | win32con.BS_PUSHBUTTON,
            150, 160, 100, 30,
            self.hwnd, 2002, wc.hInstance, None
        )

        # Set default font for better appearance
        # DEFAULT_GUI_FONT = 17
        font = win32gui.GetStockObject(17)
        win32gui.SendMessage(label1_hwnd, win32con.WM_SETFONT, font, True)
        win32gui.SendMessage(self.edit_hwnd, win32con.WM_SETFONT, font, True)
        win32gui.SendMessage(label2_hwnd, win32con.WM_SETFONT, font, True)
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
            if control_id == 2002:  # Continue button ID
                self._on_continue()
                return 0
        elif msg == win32con.WM_CLOSE:
            self._on_close()
            return 0
        elif msg == win32con.WM_DESTROY:
            win32gui.PostQuitMessage(0)
            return 0

        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def _on_continue(self):
        """Handle Continue button click"""
        print("Continue clicked - Moving to next replicate/sample")
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
    window = ReplicatePumpWindow()

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

                has_continue_button = False
                has_correct_text = False
                edit_count = 0

                for child in win.iter_children():
                    class_name = child.friendly_class_name()
                    text = child.window_text()
                    print(f"  - Class: '{class_name}' | Text: '{text}'")

                    if class_name == 'Button' and 'continue' in text.lower():
                        has_continue_button = True
                    if class_name == 'Edit':
                        edit_count += 1
                        if "Pull sample through flow cell" in text:
                            has_correct_text = True

                print(f"\n{'='*50}")
                print(f"✓ Found {edit_count} Edit control(s)")
                print(f"✓ Has Continue button: {has_continue_button}")
                print(f"✓ Has correct text: {has_correct_text}")

                if has_continue_button and has_correct_text and edit_count > 0:
                    print("✓✓✓ SUCCESS: Window is fully compatible with autopump_v2_0.py")
                    print("✓✓✓ is_correct_operator_request() will detect this window")
                    print("✓✓✓ click_continue_button() will work")
                else:
                    print("✗ WARNING: Some features not detected correctly")
                print(f"{'='*50}\n")
                break
    except Exception as e:
        print(f"Test error: {e}")

    # Keep main thread alive to see the window
    print("Window is running. Click 'Continue' button or close the window to exit.")
    try:
        while window.running:
            time.sleep(0.1)
    except KeyboardInterrupt:
        print("\nExiting...")
        window._on_close()