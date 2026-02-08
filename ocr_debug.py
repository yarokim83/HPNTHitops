"""
Diagnostic tool for OCR recognition (Windows Vision API)
Reads debug_smart_nav.png and prints all detected text.
"""
import ocr_helpers
from PIL import Image
import os

def debug_ocr():
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    debug_img_path = os.path.join(assets_dir, 'debug_smart_nav.png')
    
    if not os.path.exists(debug_img_path):
        print(f"Error: Debug image not found at {debug_img_path}")
        print("Run main.py first to generate the screenshot.")
        return

    print(f"Analyzing: {debug_img_path}")
    try:
        img = Image.open(debug_img_path)
        print(f"Image Size: {img.size}")
        
        # Initialize OCR
        ocr_helpers.init_ocr()
        
        print("\n--- Testing with ocr_helpers (Windows Vision or Tesseract) ---")
        keywords = ["Maintenance", "Repair", "Inventory", "Purchase", "Request", "M&R", "Protection", "Planning", "Administration"]
        found_any = False
        
        for kw in keywords:
            box = ocr_helpers.find_text_in_image(img, kw)
            if box:
                print(f"✅ FOUND: '{kw}' at ({box.left}, {box.top}, {box.width}, {box.height})")
                found_any = True
                
        if found_any:
            print("\n🎉 SUCCESS: OCR is working!")
        else:
            print("\n❌ FAILURE: Could not find any target keywords.")
            
    except Exception as e:
        print(f"Error during OCR analysis: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    debug_ocr()
