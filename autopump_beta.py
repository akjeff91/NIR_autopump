import time
import serial
from pywinauto import Desktop
import keyboard
import threading

SERIAL_PORT = "COM5"
runtime_sec = 60
time.sleep(2) # Wait for initialization
window_name = "Operator Request"

def pump_sample():
    with serial.Serial(SERIAL_PORT, 1200, timeout=1) as ser: # calculate baud rate for each unique connection
        ser.write(f"START:{runtime_sec}\n".encode())
        print(f"Command sent: START:{runtime_sec}")
        while True:
            line = ser.readline().decode().strip()
            if line:
                print("Pico says:", line)
            if line == "Pump stopped":
                break
            time.sleep(0.5)

def click_continue_button(window):
    try:
        for ctrl in window.descendants():
            if ctrl.friendly_class_name() == 'Button':
                ctrl.click_input()  # or .invoke() if needed
                print("Clicked 'Continue'")
        else:
            print("Continue button not available")
    except Exception as exc:
        print(f"Failed to click Continue: {exc}")
    
def is_correct_operator_request(window):
    try:
        for ctrl in window.descendants():
            if ctrl.friendly_class_name() == 'Edit':
                text = ctrl.window_text().strip()
                if "Pull sample through flow cell" in text:
                    return True
        return False
    except Exception as exc:
        print(exc)
        return False

def listen_for_f8():
    cooldown_sec = runtime_sec + 1
    last_trigger_time = 0
    while True:
        keyboard.wait("f8")
        now = time.time()
        if now - last_trigger_time >= cooldown_sec:
            print("F8 pressed - starting manual injection")
            last_trigger_time = now
            pump_sample()
        else:
            print(f"F8 pressed too soon. Wait {cooldown_sec - (now - last_trigger_time):.1f} more seconds.")

if __name__ == "__main__":
    threading.Thread(target=listen_for_f8, daemon=True).start()
    while True:
        time.sleep(1)
        try:
            windows = Desktop(backend="win32").windows()
            for win in windows:
                if window_name in win.window_text():
                    print(f"Found window titled: {win.window_text()}")
                    if is_correct_operator_request(win):
                        print("Correct 'Operator Request' detected!")
                        pump_sample()
                        click_continue_button(win)
                        time.sleep(3)
                        break  # prevent double triggering
        except Exception as e:
            print("Error:", e)