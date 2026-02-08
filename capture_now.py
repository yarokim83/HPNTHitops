"""
Capture the current screen immediately to get a valid image of the visible HI-TOPS window.
"""
from PIL import ImageGrab
import os
import time

# Wait a bit to ensure user context is stable
time.sleep(1)

# Capture full screen (all monitors)
screenshot = ImageGrab.grab(all_screens=True)

# Save
assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
if not os.path.exists(assets_dir):
    os.makedirs(assets_dir)
    
output_path = os.path.join(assets_dir, 'manual_capture.png')
screenshot.save(output_path)

print(f"Captured screen to {output_path}")
print(f"Size: {screenshot.size}")
