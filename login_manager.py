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
                 return hwnd
    return None

def perform_login(password):
    """
    Automates the login process by focusing the window, clearing the field, and pasting the password.
    
    Args:
        password (str): The password to login with.
    """
    print("Waiting for application to load (5 seconds)...")
    time.sleep(5)  # Increased wait time just in case

    # Attempt to find the window
    # Based on screenshot, title might be "Login" or "HITOPS" or "HITOPSIII"
    target_titles = ["Login", "HITOPS", "Hitops3", "HPNT"]
    hwnd = get_app_window(target_titles)
    
    if hwnd:
        print(f"Found window: {hwnd}, bringing to front...")
        try:
            # Force restore if minimized
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            # Bring to front
            win32gui.SetForegroundWindow(hwnd)
            time.sleep(0.5) 
        except Exception as e:
            print(f"Error focusing window: {e}")
    else:
        print("Could not find application window automatically. Assuming active window is correct.")

    print("Typing password...")
    
    # 1. Clear the field (Ctrl+A -> Delete) to remove any garbage
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
