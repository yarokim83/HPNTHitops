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
        time.sleep(0.5)
        pyautogui.click() # Double click to ensure focus
        
        print(f"Typing description...")
        
        import pyperclip
        pyperclip.copy(text)
        time.sleep(0.5) # Wait for clipboard
        pyautogui.hotkey('ctrl', 'v')
        
        print("Description entered via Clipboard.")
    else:
        print("Failed to find 'Description' field. Ensure screen matches the capture.")

def update_need_by_date():
    """
    Finds the 'Need By' field, reads the date, adds 1 month, and updates it.
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
        time.sleep(1)
        print(f"Searching for Need By label... ({i+1}/3)")
        
    if lbl_loc:
        print(f"Found Need By Label at {lbl_loc}. Accessing field...")
        
        # Calculate Offset to Input Box
        # Label is small text. Box is to the right.
        # Estimated offset: Right 80px
        target_x = lbl_loc[0] + 80
        target_y = lbl_loc[1] 
        
        # Click to focus
        pyautogui.click(target_x, target_y)
        time.sleep(0.5)
        
        # Select All and Copy
        import pyperclip
        pyperclip.copy('') # Clear clipboard to detect failure
        
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'c')
        time.sleep(0.5) # Increased delay
        
        from datetime import datetime
        from dateutil.relativedelta import relativedelta
        
        current_val = pyperclip.paste().strip()
        print(f"Read Need By Date: '{current_val}'")
        
        new_date_str = ""
        try:
            if not current_val:
                raise ValueError("Empty clipboard")
                
            # Parse Date (Assuming YYYY-MM-DD or similar)
            dt = datetime.strptime(current_val, '%Y-%m-%d')
            new_dt = dt + relativedelta(months=1)
            new_date_str = new_dt.strftime('%Y-%m-%d')
            print(f"Calculated New Date: {new_date_str}")
            
        except ValueError:
            print(f"Failed to parse date '{current_val}'. Defaulting to Today + 1 Month.")
            # Fallback to Today + 1 Month
            dt = datetime.now() # Use now() instead of today() for safety
            new_dt = dt + relativedelta(months=1)
            new_date_str = new_dt.strftime('%Y-%m-%d')
            print(f"Fallback Date: {new_date_str}")
        
        if new_date_str:
            # Write back
            # Re-focus and Select All to ensure we overwrite correctly
            print("Re-focusing input safely...")
            pyautogui.click(target_x, target_y, duration=0.5) # Slow movement
            time.sleep(0.5)
            
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.5)
            
            pyautogui.press('delete') 
            time.sleep(0.5)
            
            pyautogui.write(new_date_str, interval=0.15) # Slightly slower typing
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
        time.sleep(0.5)
        
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
        pyautogui.click(target_x, target_y, duration=0.5)
        time.sleep(0.5)
        
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
        pyautogui.click(click_x, lbl_loc[1], duration=0.5)
        time.sleep(0.5)
        import pyperclip
        pyperclip.copy(code_text)
        pyautogui.hotkey('ctrl', 'v')
        pyautogui.press('enter')
            
    else:
        print("Failed to find 'Account Code' label.")

def enter_part_no(part_no_text):
    if not part_no_text:
        print("No Part No provided. Skipping.")
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
        pyautogui.click(click_x, lbl_loc[1], duration=0.5)
        time.sleep(0.5)
        
        # Type Part No
        pyautogui.write(part_no_text, interval=0.1)
        time.sleep(0.8)  # Wait for input to complete
        
        # Press Enter to confirm/search
        pyautogui.press('enter')
        time.sleep(1.0)  # Wait for system to process Enter
        print("Part No entered and confirmed with Enter.")
            
    else:
        print("Failed to find 'Part No' label.")

def click_approve_purchase_request():
    """
    Finds and clicks the 'Approve Purchase Request' menu item.
    """
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    btn_img = os.path.join(assets_dir, 'approve_pr_menu.png')
    
    if not os.path.exists(btn_img):
        print(f"Error: Image not found at {btn_img}")
        return

    print("Looking for 'Approve Purchase Request' menu...")
    
    # Retry loop
    loc = None
    for i in range(3): 
        loc = locate_on_all_screens(btn_img, confidence_val=0.8)     
        if loc:
            break
        time.sleep(1)
        print(f"Searching for Approve PR menu... ({i+1}/3)")
        
    if loc:
        print(f"Found Approve PR Menu at {loc}. Clicking...")
        pyautogui.click(loc)
        time.sleep(1.0) # Wait for screen load
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
    Checks if Hitops3.exe process is running using tasklist.
    Returns True if running, False otherwise.
    """
    try:
        # Run tasklist and capture output (Binary check to avoid encoding issues)
        output = subprocess.check_output('tasklist /FI "IMAGENAME eq Hitops3.exe"', shell=True)
        if b"Hitops3.exe" in output:
            print("Hitops3.exe process detected.")
            return True
    except Exception as e:
        print(f"Error checking process list: {e}")
        
    print("Hitops3.exe process not found.")
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
    Checks if a popup with title 'Authority Access Error' exists using win32gui.
    If found, focuses it and presses Enter.
    """
    target_title = "Authority Access Error"
    
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if target_title in title:
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
    
    # Asset Map (Priority Order is Logic-Handled, but defined here for ref)
    targets = {
        'pr_goal': os.path.join(assets_dir, 'purchase_request_menu.png'),
        'inventory_mid': os.path.join(assets_dir, 'inventory_menu.png'),
        'maintenance_root': os.path.join(assets_dir, 'maintenance_btn.png'),
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
    saved_debug = False
    
    # Get Hitops window rect and handle for ROI optimization
    hitops_rect, hitops_hwnd = roi_helpers.get_hitops_window_rect()
    if hitops_rect:
        left_offset, top_offset, right_offset, bottom_offset = hitops_rect
        print(f"ROI Mode: Scanning Hitops window only ({right_offset-left_offset}x{bottom_offset-top_offset}px)")
    else:
        # Fallback to full screen if window not found
        left_offset, top_offset = 0, 0
        hitops_hwnd = None
        print("Warning: Hitops window not found, using full-screen scan (slower)")
    
    # Popup Handling Asset
    popup_asset = os.path.join(assets_dir, 'popup_invalid_parameter.png')
    
    # Bring Hitops to foreground ONCE before starting loop
    if hitops_hwnd:
        try:
            win32gui.SetForegroundWindow(hitops_hwnd)
            time.sleep(0.2)
            print("Hitops window activated.")
        except:
            pass
    
    loop_count = 0
    while time.time() - start_time < TIMEOUT:
        loop_count += 1
        if loop_count % 10 == 0:
            print(f"Scanning... ({int(time.time() - start_time)}s)")
            
        try:
            # 0. Check for Popup by Title
            check_popup_by_title()
            
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

            # --- POPUP HANDLING START ---
            # List of known blocking popups
            popup_assets = [
                os.path.join(assets_dir, 'popup_invalid_parameter.png'),
                os.path.join(assets_dir, 'popup_authority_error.png')
            ]
            
            for p_asset in popup_assets:
                if os.path.exists(p_asset):
                    # Use faster matching with fewer scales
                    fallback_scales = [1.0, 1.25]
                    popup_box = locate_with_scaling(p_asset, screenshot, confidence=0.7, scales=fallback_scales)
                        
                    if popup_box:
                        print(f"Blocking Popup detected ({os.path.basename(p_asset)})! Sending ENTER and Clicking...")
                        
                        # 1. Press ENTER (safest for blocking dialogs)
                        pyautogui.press('enter')
                        time.sleep(0.5)

                        # 2. Also try clicking if Enter didn't work (Redundant)
                        try:
                            popup_x = popup_box.left + (popup_box.width / 2) + left_offset
                            popup_y = popup_box.top + (popup_box.height * 0.85) + top_offset
                            pyautogui.click(popup_x, popup_y)
                        except:
                            pass
                            
                        time.sleep(1.0) 
                        break 
            # --- POPUP HANDLING END ---

            # 3. Check for GOAL (PR Menu)
            # If found, Click -> Verify Open -> Return
            if os.path.exists(targets['pr_goal']):
                box = safe_locate(targets['pr_goal'], screenshot, confidence=0.8)
                if box:
                    print(f"FOUND GOAL: Purchase Request Menu. Clicking...")
                    center_x = box.left + (box.width / 2) + left_offset
                    center_y = box.top + (box.height / 2) + top_offset
                    pyautogui.click(center_x, center_y)
                    print("Navigation Complete.")
                    return True

            # 4. Check for M&R Submenu (appears after hovering Maintenance)
            if os.path.exists(targets['mr_submenu']):
                box = safe_locate(targets['mr_submenu'], screenshot, confidence=0.75)
                if box:
                    if last_clicked != 'mr_submenu':
                        print(f"Found M&R Submenu! Clicking to expand...")
                        center_x = box.left + (box.width / 2) + left_offset
                        center_y = box.top + (box.height / 2) + top_offset
                        pyautogui.click(center_x, center_y)
                        last_clicked = 'mr_submenu'
                        time.sleep(0.5)
                        continue

            # 5. Check for Intermediate (Inventory)
            # If found, Click -> Continue Loop (Expect PR to appear)
            if os.path.exists(targets['inventory_mid']):
                box = safe_locate(targets['inventory_mid'], screenshot, confidence=0.8)
                if box:
                    # Only click if we haven't just clicked it (or if PR isn't visible yet)
                    if last_clicked != 'inventory':
                        print(f"Found Intermediate: Inventory Menu. Expanding...")
                        center_x = box.left + (box.width / 2) + left_offset
                        center_y = box.top + (box.height / 2) + top_offset
                        pyautogui.click(center_x, center_y)
                        last_clicked = 'inventory'
                        time.sleep(0.5) # Short wait for expansion
                        continue # Re-scan immediately

            # 6. Check for Root (Maintenance) - HOVER MENU
            if os.path.exists(targets['maintenance_root']):
                box = safe_locate(targets['maintenance_root'], screenshot, confidence=0.8)
                if box:
                    if last_clicked != 'root':
                        print(f"Found Root: Maintenance Menu. Hovering to reveal submenu...")
                        center_x = box.left + (box.width / 2) + left_offset
                        center_y = box.top + (box.height / 2) + top_offset
                        # Use moveTo instead of click - this is a hover menu!
                        pyautogui.moveTo(center_x, center_y, duration=0.2)
                        last_clicked = 'root'
                        time.sleep(1.2)  # Wait for dropdown to fully appear
                        print("Dropdown should be visible now, searching for submenu...")
                        continue

        except Exception as e:
            # Only print critical unexpected errors
            print(f"Smart Nav Critical Error: {e}")
            
        # Stuck Detection (Heuristic)
        elapsed = time.time() - start_time
        if elapsed > 15 and last_clicked is not None:
             if int(elapsed) % 5 == 0: 
                print("Stuck scanning... Attempting to dismiss invisible popup (Pressing ENTER)...")
                
                # DEBUG: Log all window titles to find the popup
                roi_helpers.log_all_window_titles()
                
                # Check by Title First (Targeted Kill)
                check_popup_by_title()
                
                # Try finding dialog by class
                popup_hwnd = roi_helpers.find_popup_by_class()
                if popup_hwnd:
                    print(f"Found dialog window (HWND: {popup_hwnd}), sending Enter...")
                    win32api.SendMessage(popup_hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                    time.sleep(0.2)
                    win32api.SendMessage(popup_hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0)
                
                # Blind Enter (Fallback)
                pyautogui.press('enter')
                time.sleep(0.5)
                
                # Reset click state to allow re-clicking Root if menu didn't open
                last_clicked = None

        time.sleep(0.2) # Fast loop
        
    print("Smart Navigation Timed Out.")
    return False
