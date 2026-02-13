"""
Debug script to capture Monitoring submenu and find RCC position.
Hovers over Monitoring, waits, and captures the submenu area.
"""
import pyautogui
import time
import os
import win32gui
import win32api
import win32con
from PIL import ImageGrab

# Get the M&C menu item location relative to Monitoring submenu
assets_dir = os.path.join(os.path.dirname(__file__), 'assets')

# Import helpers
import ocr_helpers
import roi_helpers

ocr_helpers.init_tesseract()

print("Debug: Hovering over (971, 479) in 3 seconds...")
time.sleep(3)

# Hover at Monitoring position (from logs: 971.5, 479.5)
pyautogui.moveTo(971, 479)
time.sleep(2.0)

# Capture screenshot
print("Capturing screenshot...")
screenshot = ImageGrab.grab(all_screens=True)
screenshot.save(os.path.join(os.path.dirname(__file__), 'debug_monitoring_submenu.png'))
print("Saved debug_monitoring_submenu.png")

# Also try OCR on the full image
import pytesseract
left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
print(f"Virtual screen offset: ({left_offset}, {top_offset})")

# Run OCR on the submenu area
data = pytesseract.image_to_data(screenshot, output_type=pytesseract.Output.DICT, config='--psm 11 --oem 1')
n_boxes = len(data['text'])
print(f"\nAll OCR text found (conf >= 30):")
for i in range(n_boxes):
    conf_val = int(data['conf'][i])
    text = data['text'][i].strip()
    if conf_val >= 30 and text:
        x = data['left'][i]
        y = data['top'][i]
        w = data['width'][i]
        h = data['height'][i]
        # Only show items near the submenu area (x around 900-1000, y around 400-550)
        if 800 < x < 1100 and 350 < y < 600:
            print(f"  [{conf_val}%] '{text}' at ({x}, {y}, w={w}, h={h}) -> abs ({x+left_offset}, {y+top_offset})")
