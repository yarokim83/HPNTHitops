import pyautogui
import time
import pyperclip
import win32gui
import win32con
import roi_helpers
import pyautogui

# Global configuration
pyautogui.FAILSAFE = False

def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd):
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def get_app_window(partial_title_list):
    top_windows = []
    win32gui.EnumWindows(window_enum_handler, top_windows)
    for hwnd, title in top_windows:
        for partial_title in partial_title_list:
             if partial_title.lower() in title.lower():
                 if "outlook" in title.lower():
                     continue
                 if "explorer" in title.lower() or "파일 탐색기" in title.lower():
                     continue
                 if "everything" in title.lower():
                     continue
                 return hwnd, title
    return None, None

def perform_login(password):
    """
    Automates the login process by focusing the window, clicking center, and typing password.
    Uses Process ID (via roi_helpers) to find window reliably.
    """
    print("Waiting for application to load (Searching for 30 seconds)...")
    
    hwnd = None
    title = None
    is_login_window = False
    target_rect = None

    # Retry loop for 30 seconds
    for i in range(30):
        # 1. Try Process-based Detection (Most Reliable)
        try:
            rect, p_hwnd = roi_helpers.get_hitops_window_rect()
            if p_hwnd:
                hwnd = p_hwnd
                target_rect = rect
                title = win32gui.GetWindowText(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]
                
                # Heuristic: Login window is small (e.g. < 1000px width)
                if width < 1000:
                    is_login_window = True
                    print(f"Window size ({width}x{height}) suggests Login Screen.")
                elif "login" in title.lower():
                     is_login_window = True
                
                break
        except:
            pass
        
        time.sleep(1)
        if i % 5 == 0:
            print(f"Searching for login window... ({i}/30)")

    if not hwnd:
        print("Hitops window (Process ID) not found after 30 seconds.")
        # Fallback to title search just in case
        target_titles = ["Login", "HITOPS", "Hitops3"]
        hwnd, title = get_app_window(target_titles)
        if not hwnd:
            return False

    print(f"Target Window Found: '{title}' ({hwnd})")

    if is_login_window or "login" in title.lower():
        print("Proceeding with authentication.")
        
        # 1. Activate Window
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            win32gui.SetForegroundWindow(hwnd)
        except Exception as e:
            print(f"Window activation warning: {e}")
            pyautogui.press('alt') # Wake up
        
        time.sleep(0.1)

        # 2. Click Center to Ensure Focus
        # Refresh rect after restore to get actual coordinates
        try:
             target_rect = win32gui.GetWindowRect(hwnd)
        except Exception as e:
             print(f"Failed to refresh window rect: {e}")

        if target_rect:
             cx = target_rect[0] + (target_rect[2] - target_rect[0]) // 2
             cy = target_rect[1] + (target_rect[3] - target_rect[1]) // 2
             print(f"Clicking window center to focus: {cx}, {cy}")
             try:
                 pyautogui.click(cx, cy)
             except Exception as e:
                 print(f"Click failed: {e}")
             time.sleep(0.1)

        # 3. Type Password
        print("Typing password...")
        # Clear field (Ctrl+A only, typing overwrites) 
        # Deleted 'delete' key press to prevent accidental file deletion on Desktop
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)

        # Type Password
        pyautogui.write(password, interval=0.01)
        time.sleep(0.1)
        
        pyautogui.press('enter')
        print("Login credentials submitted.")
        
        # 4. Wait for Main Window to Load
        print("Waiting for Main Window to load...")
        for k in range(100):
            # Try broader search
            main_hwnd, main_title = get_app_window(["Maintenance", "Repair System", "HITOPS", "HPNT", "Hi-Tops", "Hyundai"]) 
            if main_hwnd:
                 # Check if title changed from Login
                 if "login" not in main_title.lower():
                     print(f"Main Window Loaded: {main_title}")
                     return True
            time.sleep(0.3)
            if k % 10 == 0:
                print(f"Waiting for Hitops Main Window... ({k}/100)")
                
        print("Warning: Main Window not detected, but assuming login might have worked.")
        return True

    else:
        print("Detected window does not appear to be the Login screen. Assuming already logged in.")
        return True
