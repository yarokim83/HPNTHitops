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
                 return hwnd, title
    return None, None

def perform_login(password):
    """
    Automates the login process by focusing the window, clearing the field, and pasting the password.
    Checks if the window title indicates a login screen.
    
    Args:
        password (str): The password to login with.
    """
    print("Waiting for application to load (5 seconds)...")
    time.sleep(5)  # Increased wait time just in case

    # Attempt to find the window
    # Search for generic identifier first to catch both Login and Main window
    target_titles = ["Login", "HITOPS", "Hitops3", "HPNT"]
    hwnd, title = get_app_window(target_titles)
    
    if hwnd:
        print(f"Found window: '{title}' ({hwnd})")
        
        # Check if it is the login window
        if "login" in title.lower():
            print("Login window detected. Proceeding with authentication.")
            try:
                # Force restore if minimized
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                # Bring to front
                win32gui.SetForegroundWindow(hwnd)
                time.sleep(0.5) 
                
                print("Typing password...")
                
                # 1. Clear the field (Ctrl+A -> Delete)
                pyautogui.hotkey('ctrl', 'a')
                time.sleep(0.1)
                pyautogui.press('delete')
                time.sleep(0.1)

                # 2. Copy and Paste
                pyperclip.copy(password)
                pyautogui.hotkey('ctrl', 'v')
                time.sleep(0.5)
                
                pyautogui.press('enter')
                print("Login credentials submitted.")
                
            except Exception as e:
                print(f"Error interacting with window: {e}")
        else:
            print("Detected window does not appear to be the Login screen. Assuming already logged in.")
            return True
            
    else:
        print("Could not find application window automatically. Assuming active window is correct or app failed to launch.")
