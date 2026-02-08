"""
Locate the Maintenance & Repair icon using OCR on the manual capture.
"""
from PIL import Image
import pytesseract
import os
import sys

# Add current dir to path to import ocr_helpers
sys.path.append(os.getcwd())
try:
    import ocr_helpers
except ImportError:
    # If standard import fails, try direct file execution approach or simple tesseract
    pass

# Setup Tesseract path manually if needed (as in ocr_helpers)
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def find_icon_coordinates():
    image_path = os.path.join('assets', 'manual_capture.png')
    if not os.path.exists(image_path):
        print("manual_capture.png not found")
        return

    img = Image.open(image_path)
    print(f"Loaded image: {img.size}")

    # Preprocessing to improve OCR
    from PIL import ImageEnhance, ImageOps
    
    # Convert to grayscale and threshold
    img_gray = ImageOps.grayscale(img)
    img_thresh = img_gray.point(lambda x: 0 if x < 140 else 255) # Binary threshold
    
    # OCR on processed image
    print("Running OCR on preprocessed image...")
    data = pytesseract.image_to_data(img_thresh, output_type=pytesseract.Output.DICT)
    
    n_boxes = len(data['text'])
    print(f"Total text blocks found: {n_boxes}")
    
    target_found = False
    ref_stats = None # Statistics & Analysis
    ref_code = None  # Code & Design
    
    for i in range(n_boxes):
        text = data['text'][i].strip()
        if not text:
            continue
            
        (x, y, w, h) = (data['left'][i], data['top'][i], data['width'][i], data['height'][i])
        
        # 1. Direct Search
        if "Maintenance" in text or "Repair" in text:
            print(f"✅ FOUND TARGET: '{text}' at ({x}, {y})")
            # Capture icon above text
            icon_left = x + w//2 - 35
            icon_top = y - 75 
            save_icon(img, icon_left, icon_top)
            target_found = True
            break
            
        # 2. Reference Search
        if "Statistics" in text:
            print(f"  Found Reference 'Statistics' at ({x}, {y})")
            ref_stats = (x, y)
        if "Code" in text:
            print(f"  Found Reference 'Code' at ({x}, {y})")
            ref_code = (x, y)

    # Fallback Calculation
    if not target_found:
        print("\nTarget text not found. Calculating from references...")
        icon_x, icon_y = None, None
        
        if ref_stats:
            # "Maintenance" is 2 slots LEFT of "Statistics"
            # Grid spacing approx 250px? 
            # Statistics is col 3, Maintenance is col 1.
            # Gap = 2 * ColumnWidth
            
            # Based on previous OCR: Statistics X=842.
            # If Maintenance is col 1, it should be around X = 842 - 500 = 342?
            # Visually check screenshot...
            # Actually, looking at capture, Maintenance is far left.
            # Let's try x = ref_stats[0] - 530 (approx 2 columns)
            
            icon_x = ref_stats[0] - 530
            icon_y = ref_stats[1] - 75 # Icon is above text
            print(f"Calculated from Statistics: ({icon_x}, {icon_y})")
            
        elif ref_code:
            # Code is Row 1, Col 2. Maintenance is Row 3, Col 1.
            # This is harder.
            pass
            
        if icon_x:
             save_icon(img, icon_x, icon_y)

def save_icon(img, x, y, size=70):
    box = (int(x), int(y), int(x+size), int(y+size))
    print(f"Cropping icon at {box}")
    crop = img.crop(box)
    crop.save(os.path.join('assets', 'repair_icon.png'))
    print(f"Saved asset to assets/repair_icon.png")

if __name__ == "__main__":
    find_icon_coordinates()
