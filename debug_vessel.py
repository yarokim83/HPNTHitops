import ocr_helpers
import pyautogui
from PIL import ImageGrab
import time
import os

import menu_navigator

def debug_vessel():
    print("Maximizing HiTOPS window for debug...")
    menu_navigator.ensure_hitops_maximized()
    time.sleep(2)

    print("Capturing screen for Vessel Debug...")
    screenshot = ImageGrab.grab(all_screens=True)
    screenshot.save("debug_vessel_raw.png")
    print("Saved debug_vessel_raw.png")

    # Ensure Tesseract is initialized
    if not ocr_helpers.init_tesseract():
        print("Tesseract initialization failed!")
        return

    print("Running OCR check...")
    try:
        import pytesseract
        
        # Manually run pytesseract to see all data
        # Use same preprocessing as ocr_helpers
        processed, scale = ocr_helpers._preprocess_extreme(screenshot, target="text")
        processed.save("debug_vessel_processed.png")
        print("Saved debug_vessel_processed.png")

        config = '--psm 11 --oem 1'
        data = pytesseract.image_to_data(processed, output_type=pytesseract.Output.DICT, config=config)
        
        found_words = []
        print("\n--- Detected Text [Confidence > 30] ---")
        for i, text in enumerate(data['text']):
            conf = int(data['conf'][i])
            if conf > 30 and text.strip():
                print(f"Text: '{text}' (Conf: {conf})")
                found_words.append(text)
                
        if "Vessel" in found_words:
            print("\nSUCCESS: 'Vessel' found in OCR output!")
        else:
            print("\nFAILURE: 'Vessel' NOT found in OCR output.")
            
            # Check for close matches
            import difflib
            matches = difflib.get_close_matches("Vessel", found_words, n=5, cutoff=0.5)
            if matches:
                 print(f"Did you mean? {matches}")

    except Exception as e:
        print(f"OCR Error: {e}")

if __name__ == "__main__":
    time.sleep(2) # Give user time to switch windows if needed
    debug_vessel()
