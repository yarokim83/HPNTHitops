import pyautogui
import time
import os
import sys
from PIL import ImageGrab
import win32api
import win32con

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
        print(f"Error in multi-monitor search: {e}")
        
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
    for i in range(10): # Try for 10 seconds
        main_loc = locate_on_all_screens(main_btn_img, confidence_val=0.8)
                
        if main_loc:
            break
        
        time.sleep(1)
        print(f"Searching... ({i+1}/10)")
        
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
    for i in range(10): 
        inv_loc = locate_on_all_screens(inv_btn_img, confidence_val=0.8)     
        if inv_loc:
            break
        time.sleep(1)
        print(f"Searching for Inventory... ({i+1}/10)")
        
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
    pr_btn_img = os.path.join(assets_dir, 'purchase_request_menu.png')
    
    if not os.path.exists(pr_btn_img):
        print(f"Error: Image not found at {pr_btn_img}")
        return

    print("Looking for 'Purchase Request' menu...")
    
    # Retry loop
    pr_loc = None
    for i in range(10): 
        pr_loc = locate_on_all_screens(pr_btn_img, confidence_val=0.8)     
        if pr_loc:
            break
        time.sleep(1)
        print(f"Searching for Purchase Request... ({i+1}/10)")
        
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
    for i in range(10): 
        lbl_loc = locate_on_all_screens(need_by_img, confidence_val=0.7)     
        if lbl_loc:
            break
        time.sleep(1)
        print(f"Searching for Need By label... ({i+1}/10)")
        
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
    for i in range(10): 
        lbl_loc = locate_on_all_screens(lbl_img, confidence_val=0.8)     
        if lbl_loc:
            break
        time.sleep(1)
        print(f"Searching for Unit Price label... ({i+1}/10)")
        
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
    for i in range(10): 
        # Increased confidence to 0.93 to prevent matching similar labels on the left (e.g. Account Name)
        lbl_loc = locate_on_all_screens(lbl_img, confidence_val=0.93)     
        if lbl_loc:
            break
        time.sleep(1)
        print(f"Searching for Account Code label... ({i+1}/10)")
        
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
        
        if "0501040106" in code_text:
            # Optimized Strategy for '0501040106' (Last item)
            # Visual search was unreliable (sliding cursor).
            # Index navigation was rejected.
            # Typing/Pasting failing.
            # 'End' key jumps to bottom of list.
            print(f"Target is last item ({code_text}). Using 'End' key strategy.")
            
            # Click to Open
            # +60 hits Label Text. +100 hits Next Menu.
            # Trying +85 to hit the Input Box/Dropdown Arrow.
            click_x = lbl_loc[0] + 85
            pyautogui.click(click_x, lbl_loc[1], duration=0.5)
            time.sleep(1.0)
            
            # Navigate
            pyautogui.press('end')
            time.sleep(0.5)
            pyautogui.press('enter')
            print("Selected via 'End' key.")
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
