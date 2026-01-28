import pyautogui
import os
import time

def capture_region(filename):
    print(f"\n--- Capturing: {filename} ---")
    print("1. Move your mouse to the TOP-LEFT corner of the element you want to capture.")
    print("   (You have 5 seconds)")
    for i in range(5, 0, -1):
        print(f"{i}...", end=" ", flush=True)
        time.sleep(1)
    x1, y1 = pyautogui.position()
    print(f"\nTop-Left Set at: ({x1}, {y1})")
    
    print("\n2. Move your mouse to the BOTTOM-RIGHT corner of the element.")
    print("   (You have 5 seconds)")
    for i in range(5, 0, -1):
        print(f"{i}...", end=" ", flush=True)
        time.sleep(1)
    x2, y2 = pyautogui.position()
    print(f"\nBottom-Right Set at: ({x2}, {y2})")
    
    # Calculate width and height
    width = x2 - x1
    height = y2 - y1
    
    if width <= 0 or height <= 0:
        print("Error: Invalid region selection (Width/Height <= 0). Please try again.")
        return False
        
    print(f"Capturing region: X={x1}, Y={y1}, W={width}, H={height}")
    try:
        # Define assets path explicitly
        assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
        save_path = os.path.join(assets_dir, filename)
        
        pyautogui.screenshot(save_path, region=(x1, y1, width, height))
        print(f"Saved to: {save_path}")
        return True
    except Exception as e:
        print(f"Failed to save screenshot: {e}")
        return False

def main():
    if not os.path.exists('assets'):
        os.makedirs('assets')
        
    print("=== Image Capture Tool ===")
    print("Make sure the application window is OPEN and VISIBLE.")
    
    # Capture 1: Maintenance Button
    input("\n[Step 1] Press Enter to capture 'Maintenance & Repair' button...")
    capture_region('maintenance_btn.png')
    
    # Capture 2: M&R Submenu
    print("\n[Step 2] Now, please HOVER over the 'Maintenance & Repair' button so the submenu appears.")
    input("When the 'M&R' submenu is visible, Press Enter to capture it...")
    capture_region('mr_submenu_btn.png')
    
    print("\n=== Capture Complete ===")

if __name__ == "__main__":
    main()
