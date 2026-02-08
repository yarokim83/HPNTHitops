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
        # Check if window exists (not destroyed) - works even for minimized windows
        if win32gui.IsWindow(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # Match main window
            if "HiTOPS3" in title or "HI-TOPS" in title:
                # Exclude file explorer and Updater
                if "explorer" in title.lower() or "파일 탐색기" in title.lower() or "update" in title.lower():
                    return
                
                # For minimized windows, rect may be invalid, but hwnd is still valid
                rect = win32gui.GetWindowRect(hwnd)
                width = rect[2] - rect[0]
                height = rect[3] - rect[1]
                
                # If window is minimized, rect may be small (-32000, -32000, ...)
                # We still want to return the hwnd for restoration
                if width > 0 and height > 0:
                    target_rect = rect
                    target_hwnd = hwnd
                elif win32gui.IsIconic(hwnd):  # Window is minimized
                    # Return hwnd even if rect is invalid - caller can restore it
                    target_rect = None  # Invalid rect for minimized window
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
