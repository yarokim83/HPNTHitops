"""
Extract Repair icon from debug screenshot for use as image matching asset.
"""
from PIL import Image
import os

# Load the manual capture
screenshot_path = os.path.join('assets', 'manual_capture.png')
output_path = os.path.join('assets', 'repair_icon.png')

img = Image.open(screenshot_path)
print(f"Loaded screenshot: {img.size}")

# Estimated coordinates based on visual inspection of manual_capture.png
# The icon is in the 3rd row, 1st column of the tile grid
left = 435
top = 610
right = 505
bottom = 680

repair_icon = img.crop((left, top, right, bottom))
repair_icon.save(output_path)

print(f"Saved Repair icon to: {output_path}")
print(f"Icon size: {repair_icon.size}")

# Display for verification
repair_icon.show()
