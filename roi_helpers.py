"""
ROI (Region of Interest) and Debug Helpers for PRMaker
Provides window detection and debug utilities for optimized scanning.
"""
import win32gui
import win32con

def get_hitops_window_rect():
    """
    Finds the Hitops3 window and returns (rect, hwnd) tuple.
    rect = (left, top, right, bottom) or None (may be invalid if minimized)
    hwnd = window handle or None
    
    Note: This function can find minimized windows. The caller should
    use ShowWindow(SW_RESTORE) to restore minimized windows.
    """
    target_rect = None
    target_hwnd = None
    
    def enum_handler(hwnd, _):
        nonlocal target_rect, target_hwnd
        
        # Stop if we already found the best candidate (Top-most visible)
        if target_hwnd is not None:
            return

        # Check if window exists and is VISIBLE
        if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # Match main window or M&C sub-window
            if "HiTOPS3" in title or "HI-TOPS" in title or "Monitoring & Control" in title:
                # Exclude file explorer and Updater
                if "explorer" in title.lower() or "파일 탐색기" in title.lower() or "update" in title.lower():
                    return
                
                # Found a visible match!
                rect = win32gui.GetWindowRect(hwnd)
                target_rect = rect
                target_hwnd = hwnd
                    
    try:
        win32gui.EnumWindows(enum_handler, None)
    except:
        pass
        
    return target_rect, target_hwnd

def get_mc_window_rect():
    """
    Finds ONLY the 'Monitoring & Control' window.
    Returns (rect, hwnd).
    """
    target_rect = None
    target_hwnd = None
    
    def enum_handler(hwnd, _):
        nonlocal target_rect, target_hwnd
        
        if target_hwnd is not None:
            return

        if win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "Monitoring & Control" in title:
                rect = win32gui.GetWindowRect(hwnd)
                target_rect = rect
                target_hwnd = hwnd
                    
    try:
        win32gui.EnumWindows(enum_handler, None)
    except:
        pass
        
    return target_rect, target_hwnd

def log_all_window_titles():
    """
    Debug helper to print all visible window titles to console.
    Useful for identifying popup window titles that may not be detectable by other means.
    """
    print("\n=== DEBUG: Visible Windows ===")
    
    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title.strip():  # Only print non-empty titles
                print(f"  HWND {hwnd}: '{title}'")
                
    try:
        win32gui.EnumWindows(enum_handler, None)
    except:
        pass
        
    print("==============================\n")

def find_popup_by_class(class_name="#32770"):
    """
    Find standard Windows dialog popup by class name.
    #32770 is the class for standard dialog boxes.
    Returns HWND if found, None otherwise.
    """
    popup_hwnd = None
    
    def enum_handler(hwnd, _):
        nonlocal popup_hwnd
        if win32gui.IsWindowVisible(hwnd):
            if win32gui.GetClassName(hwnd) == class_name:
                # Found a dialog, check if it's actually visible and on top
                popup_hwnd = hwnd
                return False  # Stop enumeration
                
    try:
        win32gui.EnumWindows(enum_handler, None)
    except:
        pass
        
    return popup_hwnd
