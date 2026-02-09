import pyautogui
import time
import pyperclip
import win32gui
import win32con

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
    Automates the login process by focusing the window, clearing the field, and pasting the password.
    Checks if the window title indicates a login screen.
    
    Args:
        password (str): The password to login with.
    """
    print("Waiting for application to load (Searching for 30 seconds)...")
    
    target_titles = ["Login", "HITOPS", "Hitops3"] # Removed "HPNT" to avoid identifying Outlook
    hwnd = None
    title = None

    # Retry loop for 30 seconds
    for i in range(30):
        hwnd, title = get_app_window(target_titles)
        if hwnd:
            break
        time.sleep(1)
        if i % 5 == 0:
            print(f"Searching for login window... ({i}/30)")

    if not hwnd:
        print(f"Window '{target_titles}' not found after 30 seconds.")
        # DEBUG: List all windows
        print("--- Visible Windows (Debug) ---")
        def list_window_titles(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                t = win32gui.GetWindowText(hwnd)
                if t: print(f" - '{t}'")
        win32gui.EnumWindows(list_window_titles, None)
        print("-------------------------------")
        return False

    if hwnd:
        print(f"Found window: '{title}' ({hwnd})")
        
        # Check if it is the login window
        if "login" in title.lower():
            print("Login window detected. Proceeding with authentication.")
            try:
                # Force restore if minimized
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # Bring to front
                try:
                    win32gui.SetForegroundWindow(hwnd)
                except Exception as e:
                    print(f"SetForegroundWindow Warning: {e}. Trying Press Alt...")
                    pyautogui.press('alt') # Wake up
                    
                time.sleep(0.5) 
                
                print("Typing password...")
                
                # 1. Clear the field (Ctrl+A -> Delete)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.press('delete')
                time.sleep(0.1)

                # 2. Type Password directly (More reliable than Paste)
                # pyautogui.write handles special chars fine usually
                pyautogui.write(password, interval=0.05)
                time.sleep(0.5)
                
                pyautogui.press('enter')
                print("Login credentials submitted.")
                
                # Wait for Main Window to Load (Up to 60s)
                print("Waiting for Main Window to load...")
                for k in range(60):
                    # Try broader search
                    main_hwnd, main_title = get_app_window(["Maintenance", "Repair System", "HITOPS", "HPNT", "Hi-Tops", "Hyundai"]) 
                    if main_hwnd:
                        # Exclude Login window if "Login" is in title
                        if "Login" not in main_title:
                             print(f"Main Window loaded: '{main_title}'")
                             return True
                    
                    # Debug: Print all visible windows once to help user debugging
                    if k == 5: 
                         print("DEBUG: Listing all visible windows to find the correct title...")
                         def debug_enum(hwnd, ctx):
                             if win32gui.IsWindowVisible(hwnd):
                                 t = win32gui.GetWindowText(hwnd)
                                 if t: print(f" - Window: '{t}'")
                         win32gui.EnumWindows(debug_enum, None)

                    time.sleep(1)
                    if k % 10 == 0:
                        print(f"Waiting for Hitops Main Window... ({k}/60)")
                        
                print("Warning: Main Window not detected, but assuming login might have worked.")
                return True
                
            except Exception as e:
                print(f"Error interacting with window: {e}")
                return False
                
            except Exception as e:
                print(f"Error interacting with window: {e}")
        else:
            print("Detected window does not appear to be the Login screen. Assuming already logged in.")
            return True
            
    else:
        print("Could not find application window automatically. Assuming active window is correct or app failed to launch.")
