import pyautogui
import time
import os
import sys
import subprocess
import threading
import ctypes
from PIL import ImageGrab
import win32api
import win32con
import win32gui
import win32process
import roi_helpers
import ocr_helpers
import login_manager # Import Login Manager
from account_codes import ACCOUNT_CODE_PREFIXES, find_index_by_prefix

# Global configuration
# NOTE: Keep PyAutoGUI fail-safe ENABLED so the user can abort a runaway
# automation by moving the mouse to the top-left screen corner.
pyautogui.FAILSAFE = True

def draw_crosshair(img, x, y, color="red", label="Target"):
    """Draws a crosshair on an image for visual ROI verification."""
    from PIL import ImageDraw
    draw = ImageDraw.Draw(img)
    size = 20
    draw.line((x - size, y, x + size, y), fill=color, width=2)
    draw.line((x, y - size, x, y + size), fill=color, width=2)
    draw.text((x + 5, y + 5), label, fill=color)

def verify_and_execute_mouse(log_x, log_y, action="click", jitter=0):
    """
    Moves mouse to logical coordinates, verifies it reached there via OS,
    and performs the action. Uses win32api as fallback.
    """
    print(f"[Mouse] Targeting Log ({int(log_x)}, {int(log_y)}) -> Action: {action}")
    
    # 1. First Attempt: PyAutoGUI
    pyautogui.moveTo(log_x, log_y, duration=0.2)
    
    # OS Verification
    time.sleep(0.1)
    actual_log_x, actual_log_y = pyautogui.position()
    dist = ((log_x - actual_log_x)**2 + (log_y - actual_log_y)**2)**0.5
    
    # 2. Fallback: Win32 API if OS position mismatch
    if dist > 10:
        print(f"  [Position Verify] MISMATCH! OS reports ({actual_log_x}, {actual_log_y}). Retrying via Win32...")
        # win32api.SetCursorPos takes SCREEN coordinates (Logical)
        win32api.SetCursorPos((int(log_x), int(log_y)))
        time.sleep(0.1)
        actual_log_x, actual_log_y = pyautogui.position()
        dist = ((log_x - actual_log_x)**2 + (log_y - actual_log_y)**2)**0.5
    
    if dist <= 10:
        print(f"  [Position Verify] SUCCESS (Actual: {actual_log_x}, {actual_log_y})")
    else:
        print(f"  [Position Verify] FAILED! Distance: {int(dist)}px")

    # 3. Perform Action
    if action == "click":
        pyautogui.click()
        print("  [Action] Click performed.")
    elif action == "hover":
        if jitter > 0:
            print(f"  [Action] Hover with {jitter}px jitter...")
            for _ in range(2):
                pyautogui.moveRel(jitter, jitter, duration=0.1)
                pyautogui.moveRel(-jitter, -jitter, duration=0.1)
        else:
            print("  [Action] Hover static.")
    
    return actual_log_x, actual_log_y

def force_activate_window(hwnd):
    """
    Force-activate a window even when SetForegroundWindow alone fails.
    Uses AttachThreadInput trick to bypass Windows foreground restrictions.
    """
    try:
        # Step 1: If minimized, restore it
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.3)
        
        # Step 2: Get thread IDs
        foreground_hwnd = win32gui.GetForegroundWindow()
        foreground_thread_id = win32process.GetWindowThreadProcessId(foreground_hwnd)[0]
        target_thread_id = win32process.GetWindowThreadProcessId(hwnd)[0]
        
        # Step 3: Attach to foreground thread, activate, then detach
        attached = False
        if foreground_thread_id != target_thread_id:
            ctypes.windll.user32.AttachThreadInput(foreground_thread_id, target_thread_id, True)
            attached = True
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
            win32gui.BringWindowToTop(hwnd)
            win32gui.SetForegroundWindow(hwnd)
        finally:
            if attached:
                ctypes.windll.user32.AttachThreadInput(foreground_thread_id, target_thread_id, False)

        time.sleep(0.3)
        
        # Verify activation
        current_fg = win32gui.GetForegroundWindow()
        if current_fg == hwnd:
            print(f"Window activated successfully (HWND: {hwnd})")
            return True
        else:
            print(f"Warning: Foreground is {current_fg}, not {hwnd}. Trying Alt trick...")
            # Fallback: Alt key trick
            pyautogui.press('alt')
            time.sleep(0.1)
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.3)
            return True
    except Exception as e:
        print(f"force_activate_window failed: {e}")
        return False

# ============================================================
# Shared Functions (Used by both Purchase and M&C flows)
# ============================================================

def _get_config_path():
    """Returns a stable config.json path under AppData\\Local\\PRMaker.
    This avoids path inconsistency when running as a PyInstaller EXE,
    where __file__ points to a temporary _MEIPASS folder that is deleted on exit."""
    app_data = os.getenv('LOCALAPPDATA', os.path.expanduser('~'))
    config_dir = os.path.join(app_data, 'PRMaker')
    os.makedirs(config_dir, exist_ok=True)
    return os.path.join(config_dir, 'config.json')

def get_password():
    """Read password from config.json (AppData), or return default."""
    import json
    config_path = _get_config_path()
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config.get('password', 'fdjk213!@')
    except FileNotFoundError:
        return 'fdjk213!@'
    except (OSError, ValueError) as e:
        print(f"[Config] Failed to read password ({config_path}): {e}")
        return 'fdjk213!@'

def save_password(new_password):
    """Save password to config.json (AppData)."""
    import json
    config_path = _get_config_path()
    config = {}
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        pass
    except (OSError, ValueError) as e:
        print(f"[Config] Existing config unreadable, overwriting ({config_path}): {e}")
    config['password'] = new_password
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
        print(f"[Config] Password saved to: {config_path}")
    except OSError as e:
        print(f"[Config] FAILED to save password ({config_path}): {e}")

def ensure_app_ready():
    """
    Common pre-processing: Launch → Login → Maximize → Foreground.
    Used by both run_mc_sequence() and main.py's run_automation().
    Returns True if app is ready, False on failure.
    """
    exe_path = r"C:\Program Files (x86)\Hyundai-UNI\HITOPSIII\Hitops3.exe"
    password = get_password()
    
    # Step 1: Launch (if not running)
    if not is_hitops_running():
        if os.path.exists(exe_path):
            print(f"Launching {exe_path}...")
            subprocess.Popen(exe_path, cwd=os.path.dirname(exe_path))
            time.sleep(3)
        else:
            print(f"Error: Executable not found at {exe_path}")
            return False
    else:
        print("Hitops is already running.")
    
    # Step 2: Login
    print("Performing Login...")
    if not login_manager.perform_login(password):
        print("Login failed or timed out.")
        return False
    print("Login successful.")
    
    # Step 3: Maximize & Foreground
    ensure_hitops_maximized()
    
    return True

def run_mc_sequence():
    """
    Executes the M&C automation sequence.
    1. ensure_app_ready() (Launch + Login + Maximize)
    2. Hover 'Monitoring'
    3. Click 'M&C'
    4. Click 'Vessel'
    5. Click 'Berthing Schedule'
    """
    print("Starting M&C Automation Sequence...")
    
    # Check if M&C window is already open
    _, mc_hwnd = roi_helpers.get_mc_window_rect()
    if mc_hwnd:
        print(f"Monitoring & Control window detected (HWND: {mc_hwnd}). Skipping initial navigation.")
        try:
            # Maximize and bring to front
            win32gui.ShowWindow(mc_hwnd, win32con.SW_MAXIMIZE)
            win32gui.SetForegroundWindow(mc_hwnd)
            time.sleep(0.2)
            
            # Click center of M&C window to ensure true input focus
            rect = win32gui.GetWindowRect(mc_hwnd)
            cx = (rect[0] + rect[2]) // 2
            cy = (rect[1] + rect[3]) // 2
            pyautogui.click(cx, cy)
            time.sleep(0.2)
            print(f"M&C window maximized and focused (clicked center at {cx}, {cy}).")
        except Exception as e:
            print(f"Warning: Could not activate M&C window: {e}")
        
        # Proceed to searching for Vessel menu
        goto_vessel = True
    else:
        print("Monitoring & Control window not found. Performing full navigation sequence...")
        # Step 0: Common Launch/Login/Maximize
        if not ensure_app_ready():
            print("App initialization failed. Aborting M&C sequence.")
            return False
        goto_vessel = False

    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    
    if not goto_vessel:
        print("Searching for Monitoring Menu...")
        monitoring_img = os.path.join(assets_dir, 'monitoring_menu.png')
        mc_item_img = os.path.join(assets_dir, 'mc_menu_item.png')
        
        # Step 2: Find and Hover Monitoring (with Retry and OCR fallback)
        loc_monitoring = None
        for i in range(10): # Try for 10 seconds
            # 1. Image Search
            loc_monitoring = locate_on_all_screens(monitoring_img, confidence_val=0.7)
            
            # 2. OCR Fallback (If image search fails)
            if not loc_monitoring:
                screenshot = ImageGrab.grab(all_screens=True)
                res = ocr_helpers.find_text_in_image(screenshot, "Monitoring")
                if res:
                    # Add virtual screen offset
                    left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                    top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
                    # Calculate center from Box object
                    center_x = res.left + (res.width / 2) + left_offset
                    center_y = res.top + (res.height / 2) + top_offset
                    loc_monitoring = (center_x, center_y)

            if loc_monitoring:
                break
                
            time.sleep(1)
            print(f"Searching for Monitoring... ({i+1}/10)")

        if loc_monitoring:
            print(f"Hovering over Monitoring at {loc_monitoring}...")
            pyautogui.moveTo(loc_monitoring)
            time.sleep(0.3) # Wait for submenu
            
            # Step 3: Click M&C
            loc_mc = locate_on_all_screens(mc_item_img, confidence_val=0.7)
            # Detailed retry for M&C item too
            if not loc_mc:
                 for j in range(5):
                     time.sleep(0.5)
                     loc_mc = locate_on_all_screens(mc_item_img, confidence_val=0.7)
                     if loc_mc: break

            if loc_mc:
                print(f"Clicking M&C Menu Item at {loc_mc}...")
                pyautogui.click(loc_mc)
                print("M&C clicked. Waiting for window...")
                time.sleep(0.5)
                
                # Dismiss any popup that may appear after M&C click
                print("Dismissing any popup (pressing Enter)...")
                pyautogui.press('enter')
                time.sleep(0.2)
                
                # Wait for M&C window to actually appear
                mc_check = None
                for wait_i in range(15):
                    _, mc_check = roi_helpers.get_mc_window_rect()
                    if mc_check:
                        print(f"M&C window detected (HWND: {mc_check})")
                        break
                    time.sleep(1)
                    print(f"Waiting for M&C window... ({wait_i+1}/15)")
                
                if not mc_check:
                    print("M&C window did not appear after popup dismissal.")
                    return False
            else:
                print("M&C menu item not found.")
                return False
        else:
            print("Monitoring menu not found.")
            return False
    
    # Step 4: Activate M&C window and open Vessel menu
    _, mc_hwnd_now = roi_helpers.get_mc_window_rect()
    if mc_hwnd_now:
        try:
            win32gui.ShowWindow(mc_hwnd_now, win32con.SW_MAXIMIZE)
            time.sleep(0.2)
            force_activate_window(mc_hwnd_now)
            time.sleep(0.3)
            print(f"M&C window (HWND: {mc_hwnd_now}) activated and brought to front.")
        except Exception as e:
            print(f"Warning: Could not activate M&C window: {e}")
    else:
        print("Warning: M&C window not found before Alt+V. Proceeding anyway...")
    
    print("Step 4: Opening Vessel menu (Alt+V)...")
    pyautogui.hotkey('alt', 'v')
    time.sleep(0.5)
    print("Alt+V sent. Vessel menu should be open.")

    # Step 5: Select "Berthing Schedule" by image match (robust to menu reorder).
    # Falls back to Down x9 + Enter only if the image cannot be located.
    berthing_img = os.path.join(assets_dir, 'berthing_schedule.png')
    bs_loc = None
    if os.path.exists(berthing_img):
        for k in range(8):
            bs_loc = locate_on_all_screens(berthing_img, confidence_val=0.75)
            if bs_loc:
                break
            time.sleep(0.3)

    if bs_loc:
        print(f"Step 5: Clicking Berthing Schedule at {bs_loc}...")
        pyautogui.click(bs_loc)
    else:
        print("Step 5: Berthing Schedule image not found; falling back to Down x9 + Enter.")
        for _ in range(9):
            pyautogui.press('down')
            time.sleep(0.1)
        pyautogui.press('enter')
    print("Berthing Schedule selected.")

def click_rcc_menu():
    """
    RCC Automation Sequence:
    1. Launch/Login/Maximize HI-TOPS
    2. Hover 'Monitoring' to open submenu
    3. Click 'RCC' (located below M&C in the submenu)
    """
    print("Starting RCC Automation Sequence...")
    
    # Step 0: Common Launch/Login/Maximize
    if not ensure_app_ready():
        print("App initialization failed. Aborting RCC sequence.")
        return False
    
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    
    # Step 1: Find and Hover Monitoring
    print("Searching for Monitoring Menu...")
    monitoring_img = os.path.join(assets_dir, 'monitoring_menu.png')
    
    loc_monitoring = None
    for i in range(10):
        loc_monitoring = locate_on_all_screens(monitoring_img, confidence_val=0.7)
        
        if not loc_monitoring:
            screenshot = ImageGrab.grab(all_screens=True)
            res = ocr_helpers.find_text_in_image(screenshot, "Monitoring")
            if res:
                left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
                center_x = res.left + (res.width / 2) + left_offset
                center_y = res.top + (res.height / 2) + top_offset
                loc_monitoring = (center_x, center_y)
        
        if loc_monitoring:
            break
        time.sleep(1)
        print(f"Searching for Monitoring... ({i+1}/10)")
    
    if not loc_monitoring:
        print("Monitoring menu not found.")
        return False
    
    print(f"Hovering over Monitoring at {loc_monitoring}...")
    pyautogui.moveTo(loc_monitoring)
    time.sleep(0.3)
    
    # Step 2: Find M&C menu item as anchor, then click RCC below it
    mc_item_img = os.path.join(assets_dir, 'mc_menu_item.png')
    
    print("Searching for M&C menu item as anchor for RCC...")
    loc_mc = None
    for j in range(5):
        loc_mc = locate_on_all_screens(mc_item_img, confidence_val=0.7)
        if loc_mc:
            break
        time.sleep(0.5)
        print(f"Searching for M&C anchor... ({j+1}/5)")
    
    if loc_mc:
        # RCC is directly below M&C in the submenu
        # Use offset of 40px down from M&C center
        rcc_x = loc_mc[0]
        rcc_y = loc_mc[1] + 40
        print(f"M&C found at {loc_mc}. Clicking RCC at ({rcc_x}, {rcc_y}) [M&C + 40px]...")
        pyautogui.click(rcc_x, rcc_y)
        print("RCC clicked. Waiting for window...")
        time.sleep(0.5)
        return True
    else:
        print("M&C menu item not found (cannot locate RCC).")
        return False

def ensure_hitops_maximized():
    """
    Finds the HI-TOPS window and maximizes it if not already maximized.
    Also brings it to the foreground.
    """
    hitops_rect, hitops_hwnd = roi_helpers.get_hitops_window_rect()
    if hitops_hwnd:
        try:
            # Check current size directly
            rect = win32gui.GetWindowRect(hitops_hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            print(f"Current Window Size: {width}x{height}")

            # Check maximized state via GetWindowPlacement
            placement = win32gui.GetWindowPlacement(hitops_hwnd)
            is_maximized = (placement[1] == win32con.SW_SHOWMAXIMIZED)

            if is_maximized:
                print("Window is already maximized. Skipping resize.")
            else:
                # Check minimized state via IsIconic
                if win32gui.IsIconic(hitops_hwnd):
                    win32gui.ShowWindow(hitops_hwnd, win32con.SW_RESTORE)
                    time.sleep(0.5)
                print("Maximizing window for reliable menu detection...")
                win32gui.ShowWindow(hitops_hwnd, win32con.SW_MAXIMIZE)
                time.sleep(1.5)  # Wait for animation

            # Always Bring to front
            win32gui.SetForegroundWindow(hitops_hwnd)
            time.sleep(0.5)
            print("Hitops window activated and maximized.")
        except Exception as e:
            print(f"Window maximization warning: {e}")
    else:
        print("Warning: Could not find HI-TOPS window to maximize.")

def click_mc_menu():
    """Clicks the M&C menu using mc_icon.png as asset."""
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    img_path = os.path.join(assets_dir, 'mc_icon.png')
    print(f"Clicking M&C Menu using {img_path}...")
    loc = locate_on_all_screens(img_path, confidence_val=0.7)
    if loc:
        pyautogui.click(loc)
        print("M&C Menu clicked.")
        return True
    print("M&C Menu not found.")
    return False

def locate_on_all_screens(image_path, confidence_val=0.8, return_box=False):
    """
    Locates an image on the screen, supporting multi-monitor setups.
    Captures the full virtual screen, finds the image, and calculates absolute coordinates.
    Uses multi-scale matching (via safe_locate) for DPI robustness.

    Args:
        image_path: Path to the needle image.
        confidence_val: Match confidence threshold.
        return_box: If True, returns a Box-like object with absolute
            (left, top, width, height) for the match. If False (default),
            returns the absolute center (x, y) tuple.

    Returns:
        (x, y) tuple, Box-like object, or None.
    """
    try:
        # Capture all screens
        screenshot = ImageGrab.grab(all_screens=True)

        # Use multi-scale matching (handles 100%/125%/150% DPI)
        box = safe_locate(image_path, screenshot, confidence=confidence_val)

        if box:
            # Get Virtual Screen offset (top-left of the virtual desktop)
            left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
            top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)

            abs_left = box.left + left_offset
            abs_top = box.top + top_offset

            if return_box:
                class AbsBox:
                    __slots__ = ("left", "top", "width", "height")
                    def __init__(self, l, t, w, h):
                        self.left = l
                        self.top = t
                        self.width = w
                        self.height = h
                return AbsBox(abs_left, abs_top, box.width, box.height)

            center_x = abs_left + (box.width / 2)
            center_y = abs_top + (box.height / 2)
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
    add_btn_path = os.path.join(assets_dir, 'add_btn.png')
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
    for i in range(25): 
        add_loc = locate_on_all_screens(add_btn_img, confidence_val=0.8)     
        if add_loc:
            break
        time.sleep(0.2)
        print(f"Searching for Add Button... ({i+1}/25)")
        
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
    for i in range(25): 
        # Search for the field image
        field_loc = locate_on_all_screens(desc_field_img, confidence_val=0.8)     
        if field_loc:
            break
        time.sleep(0.2)
        print(f"Searching for Description field... ({i+1}/25)")
        
    if field_loc:
        print(f"Found Field at {field_loc}. Clicking center...")
        
        # Click to the right of the label to hit the text box
        target_x = field_loc[0] + 100
        target_y = field_loc[1]
        pyautogui.click(target_x, target_y)
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
    for i in range(10): 
        lbl_loc = locate_on_all_screens(need_by_img, confidence_val=0.7)     
        if lbl_loc:
            break
        time.sleep(0.2) # Reduced sleep
        print(f"Searching for Need By label... ({i+1}/10)")
        
    if lbl_loc:
        print(f"Found Need By Label at {lbl_loc}. Accessing field...")
        
        # Calculate Offset to Input Box
        # Label is small text. Box is to the right.
        target_x = lbl_loc[0] + 80
        target_y = lbl_loc[1] 
        
        # Click to focus
        pyautogui.click(target_x, target_y)
        time.sleep(0.2)
        
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
        pyautogui.write(new_date_str, interval=0.01) # Ultra Fast typing
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
        time.sleep(0.2)
        print(f"Searching for Unit Price label... ({i+1}/10)")
        
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
    for i in range(10): 
        # Increased confidence to 0.93 to prevent matching similar labels on the left (e.g. Account Name)
        lbl_loc = locate_on_all_screens(lbl_img, confidence_val=0.93)     
        if lbl_loc:
            break
        time.sleep(0.2)
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
        pyautogui.click(target_x, target_y)
        time.sleep(0.1)
        
        # Type Code
        # Paste didn't work. Typing didn't work. Index nav rejected by user.
        # User requested: "If text not found, scroll down and find it".
        # This implies Visual Search.
        # We only have asset for '0501040106'.
        
        target_asset = None
        if "0501040106" in code_text:
            target_asset = "acc_code_0501040106.png"
        
        # Handle specific visual target (Last Item) - Keeping for legacy safety
        if "0501040106" in code_text:
             print(f"Target is last item ({code_text}). Using 'End' key strategy.")
             click_x = lbl_loc[0] + 85
             pyautogui.click(click_x, lbl_loc[1])
             time.sleep(0.3)
             pyautogui.press('end')
             time.sleep(0.1)
             pyautogui.press('enter')
             return

        # Find Index dynamically (uses single-source-of-truth ACCOUNT_CODE_PREFIXES)
        target_code_prefix = code_text.split('/')[0].strip()
        matched_index = find_index_by_prefix(code_text)
        if matched_index is None:
            print(f"Code prefix '{target_code_prefix}' not found in known list.")
        
        if matched_index is not None:
             print(f"Target found at index {matched_index} ({target_code_prefix}). Using Index Navigation.")
             
             # Click to Open Dropdown (Optimized Offset +85px)
             click_x = lbl_loc[0] + 85
             pyautogui.click(click_x, lbl_loc[1])
             time.sleep(0.3) # Wait for list to open
             
             # Navigate
             pyautogui.press('home') # Ensure start at top
             time.sleep(0.1)
             
             # Scroll down N times
             for _ in range(matched_index):
                 pyautogui.press('down')
                 time.sleep(0.02) # Ultra fast scroll
                 
             pyautogui.press('enter')
             print("Selected via Index Navigation.")
             return

        # Fallback for others (Paste)
        print(f"No specific strategy for '{code_text}'. Falling back to Paste.")
        click_x = lbl_loc[0] + 85
        pyautogui.click(click_x, lbl_loc[1])
        time.sleep(0.1)
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
    for i in range(10): 
        lbl_loc = locate_on_all_screens(lbl_img, confidence_val=0.8)     
        if lbl_loc:
            break
        time.sleep(0.2)
        print(f"Searching for Part No label... ({i+1}/3)")
        
    if lbl_loc:
        print(f"Found Part No Label at {lbl_loc} (X={lbl_loc[0]}).")
        
        # Click to focus. Offset +60px guessed.
        click_x = lbl_loc[0] + 60
        pyautogui.click(click_x, lbl_loc[1], duration=0.2)
        time.sleep(0.2)

        # Type Part No via clipboard (handles non-ASCII safely)
        try:
            import pyperclip
            pyperclip.copy(str(part_no_text))
            time.sleep(0.1)
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.05)
            pyautogui.hotkey('ctrl', 'v')
        except Exception as e:
            print(f"Clipboard paste failed ({e}); falling back to typing.")
            pyautogui.write(str(part_no_text), interval=0.02)
        time.sleep(0.3)
        print("Part No entered.")
            
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
        # Request raw Box so we can target only the TOP checkbox of the pair.
        pair_loc = locate_on_all_screens(checkbox_img, confidence_val=0.8, return_box=True)
    except Exception as e:
        print(f"Error finding checkbox pair: {repr(e)}")
        return

    if not pair_loc:
        print("No checkbox pair found. Image mismatch.")
        return

    print(f"Found Checkbox Pair at L={pair_loc.left}, T={pair_loc.top}, W={pair_loc.width}, H={pair_loc.height}.")

    # 2. Click TOP checkbox of the pair (Header)
    target_x = pair_loc.left + (pair_loc.width / 2)
    target_y = pair_loc.top + (pair_loc.height / 4)  # Top Quarter
    
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
        # 3221225786 is likely an access violation or similar Windows error code.
        # We wrap this in a safe try-except block.
        output = subprocess.check_output('tasklist /FI "IMAGENAME eq Hitops3.exe"', shell=True, stderr=subprocess.STDOUT)
        if b"Hitops3.exe" in output or b"Hitops3" in output:
            print("Hitops3.exe process detected (No window found yet).")
            return True
    except subprocess.CalledProcessError as e:
        print(f"Tasklist command failed with exit code {e.returncode}: {e.output}")
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
    error_keywords = ["Authority", "Access", "Error", "Warning", "Alert", "권한", "오류", "메시지"]
    
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
            
            # Additional Check: Internal Dialog class
            # Windows Dialog class is usually "#32770"
            class_name = win32gui.GetClassName(hwnd)
            if class_name == "#32770":
                # Only dismiss if it's small (dialog-sized) and has a title
                # This is safer than just closing every dialog
                rect = win32gui.GetWindowRect(hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                if 100 < w < 600 and 100 < h < 400 and title:
                     print(f"Internal Dialog (#32770) detected: '{title}'. Dismissing...")
                     try:
                        win32gui.SetForegroundWindow(hwnd)
                        time.sleep(0.1)
                        pyautogui.press('enter')
                     except:
                        pass

    try:
        win32gui.EnumWindows(enum_handler, None)
    except:
        pass

# --- Popup Watchdog Thread ---
popup_watchdog_active = False

def popup_watchdog_loop():
    """Background thread that continuously monitors and dismisses popups"""
    global popup_watchdog_active
    print("[Watchdog] Background monitoring loop started.")
    while popup_watchdog_active:
        try:
            # Check for popups
            check_popup_by_title()
            # Small sleep to reduce CPU usage
            time.sleep(0.5)
        except Exception as e:
            # Silent fail for watchdog to prevent crashing
            pass
    print("[Watchdog] Background monitoring loop stopped.")

def start_popup_watchdog():
    global popup_watchdog_active
    if popup_watchdog_active:
        return # Already running
    
    popup_watchdog_active = True
    thread = threading.Thread(target=popup_watchdog_loop, daemon=True)
    thread.start()
    return thread

def stop_popup_watchdog():
    global popup_watchdog_active
    popup_watchdog_active = False


def click_pr_menu():
    """
    PR Automation Sequence (Linear Logic):
    1. Launch/Login/Maximize HI-TOPS
    2. Click 'Maintenance & Repair' Tile
    3. Click 'Purchase Request' (in submenu)
    Uses verify_and_execute_mouse for robust physical execution.
    """
    print("Starting PR Automation Sequence (Linear)...")
    
    # Step 0: Common Launch/Login/Maximize
    if not ensure_app_ready():
        print("App initialization failed. Aborting PR sequence.")
        return False
    
    assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
    
    # Step 1: Find and Click Maintenance & Repair Tile
    print("Searching for Maintenance & Repair Tile...")
    repair_icon_img = os.path.join(assets_dir, 'repair_icon.png')
    
    # Get Hitops window bounds to filter OCR (avoid false positives from other monitors)
    hitops_rect, _ = roi_helpers.get_hitops_window_rect()
    hitops_x_max = hitops_rect[2] if hitops_rect else 1600  # right edge of Hitops window (logical)
    
    loc_tile = None
    for i in range(10):
        # 1. Image Search
        loc_tile = locate_on_all_screens(repair_icon_img, confidence_val=0.7)
        
        # 2. OCR Fallback (Maintenance) — restricted to Hitops window area
        if not loc_tile:
            screenshot = ImageGrab.grab(all_screens=True)
            left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
            top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
            
            res = ocr_helpers.find_text_in_image(screenshot, "Maintenance")
            
            if res:
                center_x = res.left + (res.width / 2) + left_offset
                center_y = res.top + (res.height / 2) + top_offset
                
                # Filter: must be within Hitops window x-range (not VS Code / secondary monitor)
                if center_x <= hitops_x_max:
                    print(f"  OCR Candidate (in Hitops): {center_x}, {center_y}")
                    loc_tile = (center_x, center_y)
                else:
                    print(f"  OCR Candidate rejected (x={center_x:.0f} > hitops_x_max={hitops_x_max})")
        
        if loc_tile:
            break
        time.sleep(1)
        print(f"Searching for Maintenance Tile... ({i+1}/10)")
    
    if not loc_tile:
        print("Maintenance & Repair tile not found.")
        return False
    
    print(f"Clicking Maintenance Tile at {loc_tile}...")
    
    # Debug Proof for Tile Click
    try:
        screenshot = ImageGrab.grab(all_screens=True)
        draw_crosshair(screenshot, loc_tile[0], loc_tile[1], label="Maintenance Tile")
        debug_path = os.path.join(assets_dir, 'debug_tile_click.png')
        screenshot.save(debug_path)
    except: pass

    # Use verify_and_execute_mouse for click
    verify_and_execute_mouse(loc_tile[0], loc_tile[1], action="click")
    time.sleep(2.5) # Wait for window to open
    
    # Step 2: Wait for Maintenance & Repair System Window
    print("Waiting for Maintenance & Repair System Window...")
    main_rect = None
    main_hwnd = None
    for k in range(20):
        main_rect, main_hwnd = roi_helpers.get_maintenance_window_rect()
        if main_rect:
            print(f"Maintenance Window Detected: {main_rect}")
            break
        time.sleep(1.0)
        print(f"Waiting for Maintenance Window... ({k+1}/20)")
        
    if not main_rect:
        print("Maintenance & Repair System window did not appear.")
        return False
        
    # Ensure window is active/maximized
    try:
        win32gui.ShowWindow(main_hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(main_hwnd)
        time.sleep(1.0)
    except: pass

    # Step 3: Find 'Inventory' Menu in the New Window (with retry)
    print("Searching for 'Inventory' in Maintenance Window...")
    
    inventory_img = os.path.join(assets_dir, 'inventory_menu.png')
    loc_inventory = None
    
    for _ in range(5):
        loc_inventory = locate_on_all_screens(inventory_img, confidence_val=0.8)
        if loc_inventory:
            print(f"Found 'Inventory' via image search at {loc_inventory}")
            break
        time.sleep(1.0)
    
    if loc_inventory:
        print(f"Clicking Inventory Menu at {loc_inventory}...")
        # Debug proof (best-effort; silently skipped on any error)
        try:
            screenshot.save(os.path.join(assets_dir, 'debug_inventory_click.png'))
        except Exception:
            pass

        verify_and_execute_mouse(loc_inventory[0], loc_inventory[1], action="click")
        time.sleep(1.0)

        # Step 4: Click Purchase Request
        pr_goal_img = os.path.join(assets_dir, 'purchase_request_menu.png')

        print("Searching for Purchase Request menu item...")
        loc_pr = None
        for j in range(5):
            loc_pr = locate_on_all_screens(pr_goal_img, confidence_val=0.7)
            if loc_pr:
                break
            time.sleep(0.5)
            print(f"Searching for PR item... ({j+1}/5)")

        # OCR Fallback (use logical virtual-screen offsets, like other OCR fallbacks)
        if not loc_pr:
            screenshot_full = ImageGrab.grab(all_screens=True)
            res_pr = ocr_helpers.find_text_in_image(screenshot_full, "Purchase Request")
            if res_pr:
                left_offset = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
                top_offset = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
                cx = res_pr.left + (res_pr.width / 2) + left_offset
                cy = res_pr.top + (res_pr.height / 2) + top_offset
                loc_pr = (cx, cy)

        if loc_pr:
            print(f"Clicking Purchase Request at {loc_pr}...")
            verify_and_execute_mouse(loc_pr[0], loc_pr[1], action="click")
            print("Purchase Request clicked.")
            return True
        else:
            print("Purchase Request item not found (dropdown didn't open?)")
            # Save failure screenshot
            try:
                fail_shot = ImageGrab.grab(all_screens=True)
                fail_shot.save(os.path.join(assets_dir, 'debug_pr_fail.png'))
            except: pass
            return False

    print("Inventory menu not found in window.")
    # Save failure screenshot
    try:
        screenshot.save(os.path.join(assets_dir, 'debug_pr_fail.png'))
    except: pass
    return False

def smart_navigate_to_pr():
    """
    Wrapper for Backward Compatibility.
    Now redirects to the linear 'click_pr_menu' logic.
    """
    return click_pr_menu()
