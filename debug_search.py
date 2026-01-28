import pyautogui
import os
import sys
from PIL import ImageGrab
import win32api
import win32con

def debug_search():
    print("=== Diagnostic Tool ===")
    
    # 1. Check OpenCV
    try:
        import cv2
        print(f"[OK] OpenCV is installed. Version: {cv2.__version__}")
        has_opencv = True
    except ImportError:
        print("[WARNING] OpenCV is NOT installed.")
        print("   -> Without OpenCV, image matching requires PIXEL-PERFECT accuracy.")
        print("   -> Slight color changes or rendering differences will cause failure.")
        has_opencv = False
        
    # 2. Check Assets
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    main_btn = "maintenance_btn.png"
    path = os.path.join(assets_dir, main_btn)
    
    if os.path.exists(path):
        print(f"[OK] Asset found: {path}")
    else:
        print(f"[ERROR] Asset NOT found: {path}")
        return

    # 3. Test Screen Capture (All Screens)
    print("\nCapturing screen (all monitors)...")
    try:
        screenshot = ImageGrab.grab(all_screens=True)
        print(f"[OK] Screenshot captured. Size: {screenshot.size}")
        
        # Save for manual inspection
        debug_shot = "debug_screenshot.png"
        screenshot.save(debug_shot)
        print(f"   -> Saved full screen capture to '{debug_shot}'. Please check if the button is visible in this image.")
    except Exception as e:
        print(f"[ERROR] Failed to capture screen: {e}")
        return

    # 4. Perform Search
    print("\nAttempting to locate image options...")
    
    if has_opencv:
        print("   -> Using Confidence = 0.8")
        conf = 0.8
    else:
        print("   -> Using Exact Match (High failure rate)")
        conf = None
        
    try:
        if conf:
            box = pyautogui.locate(path, screenshot, confidence=conf)
        else:
            box = pyautogui.locate(path, screenshot)
            
        if box:
            print(f"[SUCCESS] Image found at: {box}")
            
            # Helper to calculate click point
            left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
            top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
            center_x = box.left + (box.width / 2) + left_offset
            center_y = box.top + (box.height / 2) + top_offset
            print(f"   -> Calculated Click Point: ({center_x}, {center_y})")
        else:
            print("[FAILURE] Image could NOT be found in the screenshot.")
            print("   -> Suggestion: Install opencv-python or re-capture image.")
            
    except Exception as e:
        print(f"[ERROR] Search failed with exception: {e}")

if __name__ == "__main__":
    debug_search()
