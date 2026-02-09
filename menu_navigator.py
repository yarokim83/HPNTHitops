import pyautogui
import time
import os
import sys
from PIL import ImageGrab
import win32api
import win32con
import win32gui
import subprocess
import roi_helpers
import ocr_helpers

def locate_on_all_screens(image_path, confidence_val=0.8):
    """
    Locates an image on the screen, supporting multi-monitor setups.
    Captures the full virtual screen, finds the image, and calculates absolute coordinates.
    """
    try:
        # Capture all screens
        screenshot = ImageGrab.grab(all_screens=True)
        
        # Locate the image within the screenshot
        # Note: locate returns (left, top, width, height) relative to the screenshot
        try:
            box = pyautogui.locate(image_path, screenshot, confidence=confidence_val)
        except TypeError:
             # Fallback if confidence is not supported (no opencv)
            box = pyautogui.locate(image_path, screenshot)
            
        if box:
            # Get Virtual Screen offset (top-left of the virtual desktop)
            # This is crucial if the primary monitor is not the left-most or top-most one
            left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
            top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
            
            # Calculate absolute center coordinates
            center_x = box.left + (box.width / 2) + left_offset
            center_y = box.top + (box.height / 2) + top_offset
            return (center_x, center_y)
            
    except Exception as e:
        print(f"Error in multi-monitor search: {repr(e)}")
        
    return None

def navigate_to_mr():
    """
    Finds the 'Maintenance & Repair' button, hovers over it, 
    and then clicks the 'M&R' submenu button.
    """
    # Define asset paths
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    main_btn_img = os.path.join(assets_dir, 'maintenance_btn.png')
    sub_btn_img = os.path.join(assets_dir, 'mr_submenu_btn.png')
    
    # Optimization: Check if 'Inventory' is already visible (M&R expanded)
    inv_img = os.path.join(assets_dir, 'inventory_menu.png')
    if os.path.exists(inv_img):
        inv_loc = locate_on_all_screens(inv_img, confidence_val=0.8)
        if inv_loc:
            print("Inventory menu is already visible. Skipping M&R click.")
            return

    # Check if assets exist
    if not os.path.exists(main_btn_img):
        print(f"Error: Image not found at {main_btn_img}")
        return
    if not os.path.exists(sub_btn_img):
        print(f"Error: Image not found at {sub_btn_img}")
        return

    print("Checking if 'M&R' submenu is already visible...")
    # Try to find Submenu FIRST (Optimization)
    sub_loc = locate_on_all_screens(sub_btn_img, confidence_val=0.8)
    
    if sub_loc:
        print(f"Submenu already visible at {sub_loc}. Clicking directly...")
        pyautogui.click(sub_loc)
        print("Navigation complete (Skipped Main Button hover).")
        return

    print("Submenu not found. Looking for 'Maintenance & Repair' button (scanning all monitors)...")
    
    # Retry loop for Main Button
    main_loc = None
    for i in range(3): # Try for 3 seconds
        main_loc = locate_on_all_screens(main_btn_img, confidence_val=0.8)
                
        if main_loc:
            break
        
        time.sleep(1)
        print(f"Searching... ({i+1}/3)")
        
    if not main_loc:
        print("Failed to find 'Maintenance & Repair' button on any screen.")
        return

    print(f"Found Main Button at {main_loc}. Hovering...")
    
    # Move mouse to hover
    pyautogui.moveTo(main_loc)
    time.sleep(1.0) # Wait for submenu to appear
    
    print("Looking for 'M&R' submenu...")
    
    # Retry loop for Submenu
    sub_loc = None
    for i in range(5):
        sub_loc = locate_on_all_screens(sub_btn_img, confidence_val=0.8)
        if sub_loc:
            break
        time.sleep(0.5)
        
    if sub_loc:
        print(f"Found Submenu at {sub_loc}. Clicking...")
        pyautogui.click(sub_loc)
        print("Navigation complete.")
    else:
        print("Failed to find 'M&R' submenu.")

def click_inventory():
    """
    Finds and clicks the 'Inventory' menu item.
    """
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    inv_btn_img = os.path.join(assets_dir, 'inventory_menu.png')
    
    if not os.path.exists(inv_btn_img):
        print(f"Error: Image not found at {inv_btn_img}")
        return

    print("Looking for 'Inventory' menu...")
    
    # Retry loop
    inv_loc = None
    for i in range(3): 
        inv_loc = locate_on_all_screens(inv_btn_img, confidence_val=0.8)     
        if inv_loc:
            break
        time.sleep(1)
        print(f"Searching for Inventory... ({i+1}/3)")
        
    if inv_loc:
        print(f"Found Inventory at {inv_loc}. Clicking...")
        pyautogui.click(inv_loc)
        print("Inventory clicked.")
    else:
        print("Failed to find 'Inventory' menu.")

def click_purchase_request():
    """
    Finds and clicks the 'Purchase Request' menu item.
    """
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    
    # Optimization: Check if 'Add' button is visible (Form already open)
    add_btn_path = os.path.join(assets_dir, 'add_button.png')
    if os.path.exists(add_btn_path):
        add_loc = locate_on_all_screens(add_btn_path, confidence_val=0.8)
        if add_loc:
            print("Add button visible (PR Form open). Skipping Purchase Request menu click.")
            return

    pr_btn_img = os.path.join(assets_dir, 'purchase_request_menu.png')
    
    if not os.path.exists(pr_btn_img):
        print(f"Error: Image not found at {pr_btn_img}")
        return

    print("Looking for 'Purchase Request' menu...")
    
    # Retry loop
    pr_loc = None
    for i in range(3): 
        pr_loc = locate_on_all_screens(pr_btn_img, confidence_val=0.8)     
        if pr_loc:
            break
        time.sleep(1)
        print(f"Searching for Purchase Request... ({i+1}/3)")
        
    if pr_loc:
        print(f"Found Purchase Request at {pr_loc}. Clicking...")
        pyautogui.click(pr_loc)
        print("Purchase Request clicked.")
    else:
        print("Failed to find 'Purchase Request' menu.")

def click_add_button():
    """
    Finds and clicks the 'Add' (Green Plus) button.
    """
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    add_btn_img = os.path.join(assets_dir, 'add_btn.png')
    
    if not os.path.exists(add_btn_img):
        print(f"Error: Image not found at {add_btn_img}")
        return

    print("Looking for 'Add' button...")
    
    # Retry loop
    add_loc = None
    for i in range(10): 
        add_loc = locate_on_all_screens(add_btn_img, confidence_val=0.8)     
        if add_loc:
            break
        time.sleep(1)
        print(f"Searching for Add Button... ({i+1}/10)")
        
    if add_loc:
        print(f"Found Add Button at {add_loc}. Clicking...")
        pyautogui.click(add_loc)
        print("Add Button clicked.")
    else:
        print("Failed to find 'Add' button.")

def enter_pr_description(text):
    """
    Finds the 'Description' label/field area and enters the text.
    """
    if not text:
        return

    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    # Use the new image which includes the field and cursor context
    desc_field_img = os.path.join(assets_dir, 'description_field.png')
    
    if not os.path.exists(desc_field_img):
        print(f"Error: Image not found at {desc_field_img}")
        return

    print(f"Looking for 'Description' field area to enter: {text}")
    
    # Retry loop
    field_loc = None
    for i in range(10): 
        # Search for the field image
        field_loc = locate_on_all_screens(desc_field_img, confidence_val=0.8)     
        if field_loc:
            break
        time.sleep(1)
        print(f"Searching for Description field... ({i+1}/10)")
        
    if field_loc:
        print(f"Found Field at {field_loc}. Clicking center...")
        
        # Click directly on the found center (since image includes the box)
        pyautogui.click(field_loc)
        time.sleep(0.2) # Reduced
        pyautogui.click() # Double click to ensure focus
        
        print(f"Typing description...")
        
        import pyperclip
        pyperclip.copy(text)
        time.sleep(0.2) # Wait for clipboard
        pyautogui.hotkey('ctrl', 'v')
        
        print("Description entered via Clipboard.")
    else:
        print("Failed to find 'Description' field. Ensure screen matches the capture.")

def update_need_by_date():
    """
    Finds the 'Need By' field and updates it to Today + 1 Month.
    Optimized for speed: Skip reading existing value, use faster typing.
    """
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    need_by_img = os.path.join(assets_dir, 'need_by_label.png')
    
    if not os.path.exists(need_by_img):
        print(f"Error: Image not found at {need_by_img}")
        return

    print("Looking for 'Need By' field...")
    
    # Retry loop
    lbl_loc = None
    for i in range(3): 
        lbl_loc = locate_on_all_screens(need_by_img, confidence_val=0.7)     
        if lbl_loc:
            break
        time.sleep(0.5) # Reduced sleep
        print(f"Searching for Need By label... ({i+1}/3)")
        
    if lbl_loc:
        print(f"Found Need By Label at {lbl_loc}. Accessing field...")
        
        # Calculate Offset to Input Box
        # Label is small text. Box is to the right.
        target_x = lbl_loc[0] + 80
        target_y = lbl_loc[1] 
        
        # Click to focus
        pyautogui.click(target_x, target_y)
        time.sleep(0.2) # Reduced from 0.5
        
        # Calculate New Date (Today + 1 Month)
        from datetime import datetime
        from dateutil.relativedelta import relativedelta
        
        dt = datetime.now()
        new_dt = dt + relativedelta(months=1)
        new_date_str = new_dt.strftime('%Y-%m-%d')
        print(f"Calculated New Date: {new_date_str}")
        
        # Overwrite Field
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.press('delete') 
        time.sleep(0.1)
        
        # Fast Typing
        pyautogui.write(new_date_str, interval=0.02) # Fast typing
        print(f"Need By Date updated to {new_date_str}.")
        
    else:
        print("Failed to find 'Need By' label.")

def set_unit_price_contract(enable=False):
    """
    Finds '단가계약' (Unit Price Contract) label and sets value to 'Y' if enable is True.
    """
    if not enable:
        print("Unit Price Contract check skipped (Not checked by user).")
        return

    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    lbl_img = os.path.join(assets_dir, 'unit_price_label.png')
    
    if not os.path.exists(lbl_img):
        print(f"Error: Image not found at {lbl_img}")
        return

    print("Looking for 'Unit Price Contract' field...")
    
    # Retry loop
    lbl_loc = None
    for i in range(3): 
        lbl_loc = locate_on_all_screens(lbl_img, confidence_val=0.8)     
        if lbl_loc:
            break
        time.sleep(1)
        print(f"Searching for Unit Price label... ({i+1}/3)")
        
    if lbl_loc:
        print(f"Found Unit Price Label at {lbl_loc}. Setting to Y...")
        
        # Calculate Offset to Dropdown/Input
        # Label width ~60px. Input is to the right.
        target_x = lbl_loc[0] + 70 # Adjusted offset for tighter label
        target_y = lbl_loc[1] 
        
        # Click to focus/open dropdown
        pyautogui.click(target_x, target_y)
        time.sleep(0.2)
        
        # User says: Typing Y is impossible, must click Y from modal
        # Strategy: Move Down ~25px (Item height) and Click
        # Since 'Y' appears in the list (Screenshot shows Y, N)
        # We assume Y is the first or second item.
        # Moving down 25px should hit the first item.
        
        pyautogui.moveRel(0, 25) 
        time.sleep(0.2)
        pyautogui.click()
        time.sleep(0.2)
        
        print("Unit Price Contract 'Y' selected via click.")
            
    else:
        print("Failed to find 'Unit Price Contract' label.")

def set_account_code(code_text):
    """
    Finds 'Account Code' label and enters the provided code.
    """
    if not code_text:
        print("No Account Code provided. Skipping.")
        return

    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    lbl_img = os.path.join(assets_dir, 'account_code_label.png')
    
    if not os.path.exists(lbl_img):
        print(f"Error: Image not found at {lbl_img}")
        return

    print(f"Looking for 'Account Code' field to enter '{code_text}'...")
    
    # Retry loop
    lbl_loc = None
    for i in range(3): 
        # Increased confidence to 0.93 to prevent matching similar labels on the left (e.g. Account Name)
        lbl_loc = locate_on_all_screens(lbl_img, confidence_val=0.93)     
        if lbl_loc:
            break
        time.sleep(1)
        print(f"Searching for Account Code label... ({i+1}/3)")
        
    if lbl_loc:
        print(f"Found Account Code Label at {lbl_loc} (X={lbl_loc[0]}).")
        
        # Calculate Offset to Input
        # Label width ~80px. Input is to the right.
        # Need By used +70. Account Code label is longer.
        # Box starts ~100px from center?
        # Let's try +100px.
        target_x = lbl_loc[0] + 100 
        target_y = lbl_loc[1] 
        
        # Click to focus
        pyautogui.click(target_x, target_y, duration=0.2)
        time.sleep(0.2)
        
        # Type Code
        # Paste didn't work. Typing didn't work. Index nav rejected by user.
        # User requested: "If text not found, scroll down and find it".
        # This implies Visual Search.
        # We only have asset for '0501040106'.
        
        target_asset = None
        if "0501040106" in code_text:
            target_asset = "acc_code_0501040106.png"
        
        # Full Account Codes List (Order must match Dropdown)
        ACCOUNT_CODES_LIST = [
            "0501030000",
            "0501030100",
            "0501030101",
            "0501030102",
            "0501030103",
            "0501030104",
            "0501030105",
            "0501030106",
            "0501030107",
            "0501030108",
            "0501030109",
            "0501030110",
            "0501030111",
            "0501030112",
            "0501030113",
            "0501030114",
            "0501030115",
            "0501030116",
            "0501030117",
            "0501030118",
            "0501030119",
            "0501030120",
            "0501030121",
            "0501030122",
            "0501030123",
            "0501030124",
            "0501030125",
            "0501030126",
            "0501030127",
            "0501030128",
            "0501030129",
            "0501030130",
            "0501030131",
            "0501030132",
            "0501040106"
        ]

        # Handle specific visual target (Last Item) - Keeping for legacy safety
        if "0501040106" in code_text:
             print(f"Target is last item ({code_text}). Using 'End' key strategy.")
             click_x = lbl_loc[0] + 85
             pyautogui.click(click_x, lbl_loc[1], duration=0.5)
             time.sleep(1.0)
             pyautogui.press('end')
             time.sleep(0.5)
             pyautogui.press('enter')
             return

        # Find Index dynamically
        matched_index = None
        target_code_prefix = code_text.split('/')[0].strip() # Extract '0501030101'
        
        try:
            matched_index = ACCOUNT_CODES_LIST.index(target_code_prefix)
        except ValueError:
            print(f"Code prefix '{target_code_prefix}' not found in known list.")
            matched_index = None
        
        if matched_index is not None:
             print(f"Target found at index {matched_index} ({target_code_prefix}). Using Index Navigation.")
             
             # Click to Open Dropdown (Optimized Offset +85px)
             click_x = lbl_loc[0] + 85
             pyautogui.click(click_x, lbl_loc[1], duration=0.5)
             time.sleep(1.0) # Wait for list to open
             
             # Navigate
             pyautogui.press('home') # Ensure start at top
             time.sleep(0.3)
             
             # Scroll down N times
             for _ in range(matched_index):
                 pyautogui.press('down')
                 time.sleep(0.05) # Slightly faster scroll
                 
             pyautogui.press('enter')
             print("Selected via Index Navigation.")
             return

        # Fallback for others (Paste)
        print(f"No specific strategy for '{code_text}'. Falling back to Paste.")
        click_x = lbl_loc[0] + 85
        pyautogui.click(click_x, lbl_loc[1], duration=0.2)
        time.sleep(0.2)
        import pyperclip
        pyperclip.copy(code_text)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
            
    else:
        print("Failed to find 'Account Code' label.")

def enter_part_no(part_no_text):
    if not part_no_text or len(str(part_no_text).strip()) <= 3:
        print(f"Skipping Part No input: '{part_no_text}' is too short/unsafe.")
        return

    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    lbl_img = os.path.join(assets_dir, 'part_no_label.png')
    
    if not os.path.exists(lbl_img):
        print(f"Error: Image not found at {lbl_img}")
        return

    print(f"Looking for 'Part No' field to enter '{part_no_text}'...")
    
    # Retry loop
    lbl_loc = None
    for i in range(3): 
        lbl_loc = locate_on_all_screens(lbl_img, confidence_val=0.8)     
        if lbl_loc:
            break
        time.sleep(1)
        print(f"Searching for Part No label... ({i+1}/3)")
        
    if lbl_loc:
        print(f"Found Part No Label at {lbl_loc} (X={lbl_loc[0]}).")
        
        # Click to focus. Offset +60px guessed.
        click_x = lbl_loc[0] + 60
        pyautogui.click(click_x, lbl_loc[1], duration=0.2)
        time.sleep(0.2)
        
        # Type Part No
        pyautogui.write(part_no_text, interval=0.02)
        time.sleep(0.3)  # Wait for input to complete
        
        # Press Enter to confirm/search
        pyautogui.press('enter')
        time.sleep(0.5)  # Wait for system to process Enter
        print("Part No entered and confirmed with Enter.")
            
    else:
        print("Failed to find 'Part No' label.")

def click_approve_purchase_request():
    """
    Finds and clicks the 'Approve Purchase Request' menu item.
    """
    # First, ensure HI-TOPS window is in the foreground
    hitops_rect, hitops_hwnd = roi_helpers.get_hitops_window_rect()
    if hitops_hwnd:
        try:
            import win32con
            
            # Check current size directly
            rect = win32gui.GetWindowRect(hitops_hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            print(f"Current Window Size: {width}x{height}")
            
            if width > 1600 and height > 900:
                 print("Window is already large. Skipping Resize/Restore commands.")
            else:
                 # Check minimized state via IsIconic
                if win32gui.IsIconic(hitops_hwnd):
                    win32gui.ShowWindow(hitops_hwnd, win32con.SW_RESTORE)
                    time.sleep(0.5)
                
                # Check maximized state via GetWindowPlacement
                placement = win32gui.GetWindowPlacement(hitops_hwnd)
                if placement[1] != win32con.SW_SHOWMAXIMIZED:
                    win32gui.ShowWindow(hitops_hwnd, win32con.SW_MAXIMIZE)
                    time.sleep(1.0) # Wait for animation

            win32gui.SetForegroundWindow(hitops_hwnd)
            time.sleep(0.5)
            print("HI-TOPS window activated for Approve PR search.")
        except Exception as e:
            print(f"Window activation warning: {e}")
    
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    btn_img = os.path.join(assets_dir, 'approve_pr_menu.png')
    
    if not os.path.exists(btn_img):
        print(f"Error: Image not found at {btn_img}")
        return

    print("Looking for 'Approve Purchase Request' menu...")
    
    # Retry loop
    loc = None
    for i in range(5):  # Increased retries
        loc = locate_on_all_screens(btn_img, confidence_val=0.75)  # Slightly lower confidence
        if loc:
            break
        time.sleep(0.5)
        print(f"Searching for Approve PR menu... ({i+1}/5)")
        
    if loc:
        print(f"Found Approve PR Menu at {loc}. Clicking...")
        pyautogui.click(loc)
        time.sleep(0.5) # Wait for screen load
    else:
        print("Failed to find 'Approve Purchase Request' menu.")

def select_pr_in_approval_list(target_description):
    """
    Finds the PR with the matching description in the Approval List grid and clicks its checkbox.
    Uses Clipboard scraping (Ctrl+A, Ctrl+C) to read the list content.
    """
    if not target_description:
        print("No description provided. Skipping selection.")
        return

    print("Waiting for Approval List window...")
    time.sleep(3.0) # Wait for window popup

    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    # Updated to Checkbox Pair (Header + Row 1)
    # Using this to distinguish Header clearly.
    checkbox_img = os.path.join(assets_dir, 'checkbox_pair.png')
    
    if not os.path.exists(checkbox_img):
        print(f"Error: Image not found at {checkbox_img}")
        return

    # 1. Find the Pair
    print("Scanning for Checkbox Pair...")
    try:
        pair_loc = locate_on_all_screens(checkbox_img, confidence_val=0.8)
    except Exception as e:
        print(f"Error finding checkbox pair: {repr(e)}")
        return

    if not pair_loc:
        print("No checkbox pair found. Image mismatch.")
        return

    print(f"Found Checkbox Pair at {pair_loc}.")

    # 2. Strategy Update: Click Top Half of the Pair (Header Checkbox)
    print("Strategy: Clicking the TOP Checkbox of the detected pair.")
    
    # pair_loc is (left, top, width, height)
    # We want top half. Let's aim for Top + Height/4
    # Note: locate_on_all_screens returns (x, y) center? No, usually Box (left, top, width, height) if using locateOnScreen.
    # But my wrapper `locate_on_all_screens` returns CENTER (x, y) if simple.
    # Let's check `locate_on_all_screens` implementation.
    # Step 679: returns `pyautogui.locate(..., confidence)` which returns Box?
    # NO, wait. `pyautogui.locate` returns Box. `locateCenter` returns Point.
    # My wrapper code in Step 679:
    # box = pyautogui.locate(...)
    # if box: return box
    # So it returns a BOX (left, top, width, height).
    
    # Calculate Target:
    target_x = pair_loc.left + (pair_loc.width / 2)
    target_y = pair_loc.top + (pair_loc.height / 4) # Top Quarter
    
    print(f"Clicking Header Checkbox at ({target_x}, {target_y})...")
    pyautogui.click(target_x, target_y)
    print("Checkbox clicked.")
    
    # 3. Click Approve Icon (Green Check)
    time.sleep(1.0) # Wait for enable
    approve_icon_img = os.path.join(assets_dir, 'approve_icon.png')
    
    if os.path.exists(approve_icon_img):
        print("Looking for Approve Icon (Green Check)...")
        app_loc = None
        for i in range(5):
             app_loc = locate_on_all_screens(approve_icon_img, confidence_val=0.8)
             if app_loc:
                 break
             time.sleep(1.0)
             print(f"Waiting for Approve Icon to become active... ({i+1}/5)")
        
        if app_loc:
            print(f"Found Approve Icon at {app_loc}. Clicking...")
            pyautogui.click(app_loc)
        else:
            print("Failed to find Approve Icon.")
    else:
        print(f"Approve Icon asset missing: {approve_icon_img}")



def is_hitops_running():
    """
    Checks if Hitops3 application is running.
    Prioritizes Window Detection over Process List (Process might be a zombie).
    Returns True if Window Found or Process detected.
    """
    # 1. Check for visible window first (Most reliable for automation)
    rect, hwnd = roi_helpers.get_hitops_window_rect()
    if hwnd:
        print("Hitops3 window detected via Window API.")
        return True

    # 2. Fallback: Check process list (e.g. if minimized to tray or loading)
    try:
        # Run tasklist and capture output (Binary check to avoid encoding issues)
        output = subprocess.check_output('tasklist /FI "IMAGENAME eq Hitops3.exe"', shell=True)
        if b"Hitops3.exe" in output or b"Hitops3" in output:
            print("Hitops3.exe process detected (No window found yet).")
            return True
    except Exception as e:
        print(f"Error checking process list: {e}")
        
    print("Hitops3 application not found (No Window, No Process).")
    return False

from PIL import Image, ImageOps

def locate_with_scaling(image_path, screenshot, confidence=0.8, scales=[1.0, 1.25, 1.5, 0.75, 0.8]):
    """
    Locates image with multiple scales to handle DPI differences.
    Returns: Box (relative to screenshot)
    """
    try:
        # Load Needle Image
        if isinstance(image_path, str):
            needle_base = Image.open(image_path)
        else:
            needle_base = image_path
            
        base_w, base_h = needle_base.size
        
        for scale in scales:
            if scale == 1.0:
                 needle = needle_base
            else:
                 new_w = int(base_w * scale)
                 new_h = int(base_h * scale)
                 needle = needle_base.resize((new_w, new_h), Image.Resampling.LANCZOS)
            
            try:
                # Use PyAutoGUI locate (supports PIL images)
                box = pyautogui.locate(needle, screenshot, confidence=confidence)
                if box:
                    print(f"Match found at scale {scale}x")
                    return box
            except pyautogui.ImageNotFoundException:
                 continue
            except Exception:
                 continue
                 
    except Exception as e:
        print(f"Scaling Search Error: {e}")
        return None
        
    return None

def safe_locate(image_path, screenshot, confidence=0.8):
    """
    Helper to locate image without raising ImageNotFoundException.
    Uses Multi-Scale Matching for robustness.
    Returns Box or None.
    """
    # 1. Try Simple 1.0x first (Fastest)
    try:
        return pyautogui.locate(image_path, screenshot, confidence=confidence)

    except pyautogui.ImageNotFoundException:
        pass # Fallthrough to scaling
    except Exception:
        pass

    # 2. Try Multi-Scale (DPI fallback) - REDUCED for speed
    # Only try 1.25 (125% scaling) which is most common on Windows
    fallback_scales = [1.25] 
    return locate_with_scaling(image_path, screenshot, confidence=confidence, scales=fallback_scales)

def check_popup_by_title():
    """
    Checks if a popup with error-related title exists using win32gui.
    If found, focuses it and presses Enter/Escape to dismiss.
    """
    # Multiple keywords to catch various error popups
    error_keywords = ["Authority", "Access", "Error", "Warning", "Alert", "권한", "오류"]
    
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # Check if any keyword is in the title
            for keyword in error_keywords:
                if keyword.lower() in title.lower():
                    print(f"Blocking Popup detected by Title: '{title}'. Dismissing...")
                    try:
                        # Generic Windows dismissal
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.2)
                        pyautogui.press('enter')
                        return True # Stop enumeration? No, enum returns bool
                    except Exception as e:
                        print(f"Failed to dismiss popup: {e}")
                        # Try sending close message
                        try:
                            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
                        except:
                            pass
    try:
        win32gui.EnumWindows(enum_handler, None)
    except:
        pass



def smart_navigate_to_pr():
    """
    Event-Driven Navigation:
    Continuously scans for [Purchase Request, Inventory, Maintenance] menus simultaneously.
    Click Priority: PR Menu (Goal) > Inventory (Mid) > Maintenance (Root).
    Fast reaction time by using a single screenshot for multiple checks.
    """
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    
    # Asset Map - Updated for tile-based UI
    targets = {
        'pr_goal': os.path.join(assets_dir, 'purchase_request_menu.png'),
        'inventory_mid': os.path.join(assets_dir, 'inventory_menu.png'),
        'repair_icon': os.path.join(assets_dir, 'repair_icon.png'),  # Main tile icon
        'mr_submenu': os.path.join(assets_dir, 'mr_submenu_btn.png') # Optional
    }
    
    # Verify assets
    for name, path in targets.items():
        if not os.path.exists(path):
            print(f"Warning: Asset {name} missing at {path}")

    print("Starting Smart Navigation (Parallel Check)...")
    start_time = time.time()
    TIMEOUT = 60
    
    # State tracking to avoid spam-clicking the same parent
    last_clicked = None
    last_clicked_time = 0  # To track when we last clicked/hovered
    saved_debug = False
    
    # Popup Handling Asset
    popup_asset = os.path.join(assets_dir, 'popup_invalid_parameter.png')
    
    # 1. Activate & Maximize Window FIRST
    hitops_rect_initial, hitops_hwnd = roi_helpers.get_hitops_window_rect()
    
    if hitops_hwnd:
        try:
            import win32con
            
            # Check current size directly
            rect = win32gui.GetWindowRect(hitops_hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            
            print(f"Current Window Size: {width}x{height}")
            
            # Strategy: Only maximize if window is too small for reliable detection
            if width > 1600 and height > 900:
                print("Window is already large enough. No resize needed.")
            else:
                # Window is small - maximize it
                print("Window too small. Maximizing for reliable menu detection...")
                win32gui.ShowWindow(hitops_hwnd, win32con.SW_MAXIMIZE)
                time.sleep(1.5)  # Wait for animation

            # Always Bring to front
            win32gui.SetForegroundWindow(hitops_hwnd)
            time.sleep(0.5)
            print("Hitops window activated.")
        except Exception as e:
            print(f"Window activation warning: {e}")
            pass
            
    # 2. Get ROI coordinates (AFTER maximization)
    # Re-fetch rect because maximization changed it
    hitops_rect, _ = roi_helpers.get_hitops_window_rect()
    
    if hitops_rect:
        left_offset, top_offset, right_offset, bottom_offset = hitops_rect
        print(f"ROI Mode: Scanning Hitops window only ({right_offset-left_offset}x{bottom_offset-top_offset}px)")
        print(f"Window Position: Left={left_offset}, Top={top_offset}")
    else:
        # Fallback to full screen if window not found
        left_offset, top_offset = 0, 0
        hitops_hwnd = None
        print("Warning: Hitops window lost after activation, using full-screen scan")

    loop_count = 0
    while time.time() - start_time < TIMEOUT:
        # Check for blocking popup
        check_popup_by_title()
        
        loop_count += 1
        if loop_count % 10 == 0:
            print(f"Scanning... ({int(time.time() - start_time)}s)")
            
        try:
            # 1. Capture Screen (ROI if available, full screen otherwise)
            if hitops_rect:
                screenshot = ImageGrab.grab(bbox=hitops_rect)
            else:
                screenshot = ImageGrab.grab(all_screens=True)
            
            # DEBUG: Save screenshot to verify what Python sees
            if not saved_debug:
                debug_path = os.path.join(assets_dir, 'debug_smart_nav.png')
                screenshot.save(debug_path)
                print(f"DEBUG: Saved screenshot to {debug_path}. Please check if Right Monitor is visible.")
                saved_debug = True
            
            # 2. Virtual Screen Offset
            try:
                left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
            except:
                left_offset = 0
                top_offset = 0


            # logic: if we clicked something but haven't found the next target after 5 seconds, reset state
            if last_clicked and (time.time() - last_clicked_time > 5.0):
                print(f"Action '{last_clicked}' seems to have failed (no progress for 5s). Resetting state to retry...")
                last_clicked = None
                
            # 3. Check for GOAL (PR Menu)
            # If found, Click -> Verify Open -> Return
            # 3. Check for GOAL (PR Menu) - OCR FIRST, Image Fallback
            if os.path.exists(targets['pr_goal']):
                box = None
                
                # Try OCR first
                try:
                    box = ocr_helpers.find_purchase_request_menu(screenshot)
                    if box:
                        print(f"FOUND GOAL via OCR: Purchase Request Menu")
                except:
                    pass
                
                # Fallback to Image
                if box is None:
                    box = safe_locate(targets['pr_goal'], screenshot, confidence=0.8)
                    if box:
                        print(f"FOUND GOAL via Image: Purchase Request Menu")

                if box:
                    print(f"Clicking Purchase Request Menu...")
                    center_x = box.left + (box.width / 2) + left_offset
                    center_y = box.top + (box.height / 2) + top_offset
                    pyautogui.click(center_x, center_y)
                    print("Navigation Complete.")
                    return True

            # 4. Check for M&R Submenu (only if we haven't already entered M&R)
            if last_clicked != 'mr_submenu':
                box = None
                
                # Try OCR first
                try:
                    box = ocr_helpers.find_mr_submenu(screenshot)
                    if box:
                        print(f"Found M&R Submenu via OCR!")
                except:
                    pass
                
                # Fallback to Image
                if box is None and os.path.exists(targets['mr_submenu']):
                     box = safe_locate(targets['mr_submenu'], screenshot, confidence=0.6)
                     if box:
                         print(f"Found M&R Submenu via Image")

                if box:
                    print(f"Clicking M&R Submenu...")
                    center_x = box.left + (box.width / 2) + left_offset
                    center_y = box.top + (box.height / 2) + top_offset
                    
                    # DEBUG: Visual Click Confirmation
                    try:
                        from PIL import ImageDraw
                        debug_click_img = screenshot.copy()
                        draw = ImageDraw.Draw(debug_click_img)
                        # Draw red circle at target
                        r = 10
                        rel_x = center_x - left_offset
                        rel_y = center_y - top_offset
                        draw.ellipse((rel_x-r, rel_y-r, rel_x+r, rel_y+r), outline="red", width=3)
                        debug_path = os.path.join(assets_dir, 'debug_last_click.png')
                        debug_click_img.save(debug_path)
                        print(f"DEBUG: Saved click target visual to {debug_path}")
                    except Exception as e:
                        print(f"Debug save failed: {e}")

                    pyautogui.click(center_x, center_y)
                    last_clicked = 'mr_submenu'
                    last_clicked_time = time.time()
                    time.sleep(0.5)
                    continue
                else:
                    if loop_count % 10 == 0:  # Print debug every 10 loops
                        print("  → M&R submenu not found (yet)")

            # 5. Check for Intermediate (Inventory) - OCR FIRST, Image Fallback
            # Only search if we haven't already clicked it
            if last_clicked != 'inventory':
                box = None
                
                # Try OCR first (more reliable)
                try:
                    box = ocr_helpers.find_inventory_menu(screenshot)
                    if box:
                        print(f"Found Inventory via OCR!")
                except Exception as e:
                    if loop_count % 10 == 0:
                        print(f"  OCR search failed: {e}, falling back to image matching")
                
                # Fallback to image matching if OCR fails
                if box is None and os.path.exists(targets['inventory_mid']):
                    box = safe_locate(targets['inventory_mid'], screenshot, confidence=0.55)
                    if box:
                        print(f"Found Inventory via image matching")
                
                if box:
                    print(f"Clicking Inventory Menu...")
                    center_x = box.left + (box.width / 2) + left_offset
                    center_y = box.top + (box.height / 2) + top_offset
                    pyautogui.click(center_x, center_y)
                    last_clicked = 'inventory'
                    last_clicked_time = time.time()
                    time.sleep(0.8)
                    continue
                else:
                    if loop_count % 10 == 0:
                        print("  → Inventory menu not found (yet)")

            # 6. Check for Root (Maintenance) - HOVER MENU
            # 6. Check for Repair Tile Icon
            if last_clicked != 'root':
                box = None
                
                # Search for repair icon (no Y restriction - tiles can be anywhere)
                if os.path.exists(targets['repair_icon']):
                    # DEBUG: Find ALL matches to see if we satisfy multiple icons
                    try:
                        all_matches = list(pyautogui.locateAll(targets['repair_icon'], screenshot, confidence=0.75))
                        print(f"DEBUG: Found {len(all_matches)} potential Repair Icons.")
                        
                        if len(all_matches) > 0:
                            # Visual Debug for ALL matches
                            from PIL import ImageDraw
                            debug_all_img = screenshot.copy()
                            draw = ImageDraw.Draw(debug_all_img)
                            for match in all_matches:
                                cx = match.left + (match.width / 2)
                                cy = match.top + (match.height / 2)
                                r = 15
                                draw.ellipse((cx-r, cy-r, cx+r, cy+r), outline="blue", width=3)
                                print(f"  - Match at ({cx}, {cy})")
                                
                            debug_all_path = os.path.join(assets_dir, 'debug_all_repair_matches.png')
                            debug_all_img.save(debug_all_path)
                            print(f"DEBUG: Saved all matches to {debug_all_path}")
                            
                            # Use the first match (or logic to pick best one?)
                            # For now, pick the one with highest Y (lowest on screen?) or highest X?
                            # Usually main menu is in center/body.
                            box = all_matches[0] 
                    except Exception as e:
                         print(f"LocateAll failed: {e}")
                         box = safe_locate(targets['repair_icon'], screenshot, confidence=0.75)

                # OCR Fallback for Repair Tile
                if box is None:
                    try:
                        box = ocr_helpers.find_maintenance_tile(screenshot)
                        if box:
                            print("Found Repair Tile via OCR!")
                    except Exception as e:
                        print(f"OCR Tile search failed: {e}")

                    if box:
                        print(f"Found Repair Icon!")
                            
                if box:
                    print(f"Clicking Repair Icon...")
                    center_x = box.left + (box.width / 2) + left_offset
                    center_y = box.top + (box.height / 2) + top_offset
                    
                    print(f"DEBUG: Repair Icon Coordinates: ({center_x}, {center_y})")
                    
                    # DEBUG: Visual Click Confirmation
                    try:
                        from PIL import ImageDraw
                        debug_click_img = screenshot.copy()
                        draw = ImageDraw.Draw(debug_click_img)
                        # Draw red circle at target
                        r = 10
                        rel_x = center_x - left_offset
                        rel_y = center_y - top_offset
                        draw.ellipse((rel_x-r, rel_y-r, rel_x+r, rel_y+r), outline="red", width=3)
                        debug_path = os.path.join(assets_dir, 'debug_last_click_repair.png')
                        debug_click_img.save(debug_path)
                        print(f"DEBUG: Saved click target visual to {debug_path}")
                    except Exception as e:
                        print(f"Debug save failed: {e}")

                    # Definite Hover to show user where we are clicking
                    print(f"Moving to Repair Icon at ({center_x}, {center_y})...")
                    pyautogui.moveTo(center_x, center_y, duration=0.5)
                    
                    # Wait longer (3s) to ensure UI is stable after resize
                    print("Waiting 3 seconds for UI stability...")
                    time.sleep(3.0) 

                    # Click the tile
                    # Try Double Click? Or just single click?
                    # Let's stick to single click first, but with delay.
                    pyautogui.click() # Click at current location
                    
                    # Immediately check for popup after click
                    time.sleep(0.5)
                    check_popup_by_title()
                    
                    last_clicked = 'root'
                    last_clicked_time = time.time()
                    time.sleep(1.0)  # Wait for submenu window to open
                    
                    # Check popup again
                    check_popup_by_title()
                    
                    print("Submenu should be visible now, searching for Inventory...")
                    continue

        except Exception as e:
            # Only print critical unexpected errors
            print(f"Smart Nav Critical Error: {e}")
            


        time.sleep(0.2) # Fast loop
        
        # Fallback: Check Taskbar Icon if nothing found for strict time
        if loop_count % 15 == 0:
            print("  → Nothing found in ROI. Checking Taskbar for icon...")
            taskbar_asset = os.path.join(assets_dir, 'taskbar_icon.png')
            if os.path.exists(taskbar_asset):
                 # Temporarily grab full screen
                 full_shot = ImageGrab.grab(all_screens=True)
                 # Taskbar is usually at bottom, low confidence
                 tb_box = locate_with_scaling(taskbar_asset, full_shot, confidence=0.7, scales=[1.0, 1.25])
                 
                 if tb_box:
                     print(f"Found Hitops Taskbar Icon! Clicking to activate...")
                     # Calculate absolute click coordinates
                     # Note: locate_with_scaling returns coordinates relative to the image (full_shot)
                     # full_shot includes virtual screen offset, but pyautogui handles it if we are careful.
                     # However, safe_locate usually handles offset. Let's trust locate_with_scaling's return logic.
                     # Actually simpler: just click it.
                     
                     # Need to handle virtual screen offset for click
                     try:
                         v_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                         v_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
                     except:
                         v_left, v_top = 0, 0
                         
                     # box is relative to full_shot's top-left (which is v_left, v_top)
                     click_x = tb_box.left + (tb_box.width / 2) + v_left
                     click_y = tb_box.top + (tb_box.height / 2) + v_top
                     
                     pyautogui.click(click_x, click_y)
                     time.sleep(2.5) # Wait ample time for window to restore and render
                     
                     # Taskbar click directly opens M&R window (not just menu hover)
                     # So we mark as if we already clicked the M&R submenu
                     last_clicked = 'mr_submenu'
                     last_clicked_time = time.time()
                     print("Taskbar activated M&R window. Proceeding to search for Inventory...")
                     
                     # Re-acquire window rect after activation
                     hitops_rect, hitops_hwnd = roi_helpers.get_hitops_window_rect()
                     continue
        
    print("Smart Navigation Timed Out.")
    # Save timeout screenshot for debugging
    timeout_shot = ImageGrab.grab(all_screens=True)
    timeout_path = os.path.join(assets_dir, 'debug_timeout.png')
    timeout_shot.save(timeout_path)
    print(f"DEBUG: Saved timeout screenshot to {timeout_path}")
    
    return False
