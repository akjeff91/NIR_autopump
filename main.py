from machine import Pin, Timer
import time
import sys
import select

relay = Pin(15, Pin.OUT)  # Control pin connected to IN1 of relay
led = Pin("LED", Pin.OUT)

def pump_sample(duration_sec):
    print("Pump starting for", duration_sec, "seconds")
    relay.value(1)  # Close relay (Start pump)
    led.value(1)
    time.sleep(duration_sec)
    relay.value(0)  # Open relay (Stop pump)
    led.value(0)
    print("Pump stopped")

while True:
    if sys.stdin in select.select([sys.stdin], [], [], 0)[0]:
        line = sys.stdin.readline().strip()
        print("Received:", line)  # Echo command for debugging
        if line.startswith("START:"):
            try:
                runtime = int(line.split(":")[1])
                pump_sample(runtime)
            except:
                print("Invalid command format")

