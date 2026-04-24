NIR autosampler compatible with Thermo Scientific OMNIC workflows that uses a Raspberry Pi Zero microcontroller and relay control to automate a peristaltic pump. Download main.py to the Raspberry Pi to run automatically on boot and run the appropriate version of autopump via terminal. The GUI appears when it is chosen in the system tray.

--- VERSIONS ---

beta: simple version that will pump sample for 60 seconds every time the sample replicate window appears

v1.0: included GUI using PyQt6 module to create widget for monitoring program output

v1.1: acts as a sample repeater that can run one sample many times, appending a letter to the end of each spectrum name (A, B, C, etc.)

v2.0: added a sample queue which can run and track progress of a sequence of samples automatically; sample queue added to the GUI as a numbered table
