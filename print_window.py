from pywinauto import Desktop

windows = Desktop(backend="uia").windows()

for win in windows:
    try:
        if "Operator Request" in win.window_text():
            popup = Desktop(backend="uia").window(title=win.window_text())
            popup.print_control_identifiers()
            #if "Hello" in popup.child_window(control_type="Edit").window_text():
            #    print("Correct popup detected!")
    except Exception as e:
        print(f"Could not access window: {e}")