"""
ROI (Region of Interest) Helpers
Handles window detection and coordinate calculations.
Updated with UIA support for robust window finding.
"""
import win32gui
import win32con
from ctypes import windll
import win32api

def get_hitops_window_rect():
    """
    Classic Win32 API method.
    Returns (left, top, right, bottom) of the specific HiTOPS window.
    """
    hwnd = None
    # Fuzzy search
    def enum_handler(h, ctx):
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h)
            if "HITOPS" in title.upper() or "HI-TOPS" in title.upper():
                ctx.append(h)
                
    found = []
    win32gui.EnumWindows(enum_handler, found)
    if found:
        hwnd = found[0]

    if hwnd:
        try:
            rect = win32gui.GetWindowRect(hwnd)
            return rect, hwnd
        except:
            pass
            
    return None, None

def get_mc_window_rect():
    """
    Detection for Monitoring & Control (M&C) window.
    Searches for "Monitoring" or "M&C" titles.
    Ignores the main HITOPS window to prevent false positives.
    """
    _, hitops_hwnd = get_hitops_window_rect()
    
    hwnd = None
    def enum_handler(h, ctx):
        if h == hitops_hwnd:
            return
            
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h)
            if "MONITOR" in title.upper() or "M&C" in title.upper():
                ctx.append(h)
                
    found = []
    win32gui.EnumWindows(enum_handler, found)
    if found:
        hwnd = found[0]

    if hwnd:
        try:
            rect = win32gui.GetWindowRect(hwnd)
            return rect, hwnd
        except:
            pass
            
    return None, None

def get_maintenance_window_rect():
    """
    Detection for Maintenance & Repair System window.
    Searches for "Maintenance & Repair" or "Repair System" titles.
    """
    hwnd = None
    def enum_handler(h, ctx):
        if win32gui.IsWindowVisible(h):
            title = win32gui.GetWindowText(h)
            if "Maintenance" in title and "Repair" in title:
                ctx.append(h)
                
    found = []
    win32gui.EnumWindows(enum_handler, found)
    if found:
        hwnd = found[0]

    if hwnd:
        try:
            rect = win32gui.GetWindowRect(hwnd)
            return rect, hwnd
        except:
            pass
            
    return None, None

def get_hitops_window_rect_uia():
    """
    Modern UI Automation method (via pywinauto).
    More robust for .NET forms and multi-monitor setups.
    Returns (left, top, right, bottom) tuple or None.
    """
    try:
        from pywinauto import Application
        
        # Connect to existing app (timeout 3s)
        app = Application(backend="uia").connect(title_re=".*HiTOPS.*|.*HITOPS.*|.*HI-TOPS.*", timeout=3)
        win = app.window(title_re=".*HiTOPS.*|.*HITOPS.*|.*HI-TOPS.*")
        
        if win.exists():
            r = win.rectangle()
            # rectangle() from pywinauto is often PHYSICAL pixels if app is DPI aware
            # or LOGICAL if not. Let's assume it's consistent with Win32.
            return (r.left, r.top, r.right, r.bottom)
            
    except Exception as e:
        print(f"UIA Rect fetch failed: {e}")
        
    return None


def get_dpi_scaling():
    """
    Returns the system DPI scaling factor (e.g., 1.25 for 125%).
    Uses shcore for the most reliable detection on modern Windows.
    """
    try:
        import ctypes
        # DEVICE_PRIMARY = 0
        scale = ctypes.windll.shcore.GetScaleFactorForDevice(0)
        return scale / 100.0
    except:
        try:
            # Fallback to older method
            import win32api
            import win32con
            logical_width = win32api.GetSystemMetrics(win32con.SM_CXSCREEN)
            # This fallback might return 1.0 if process is already DPI-aware
            # but it's better than nothing.
            return 1.0 # Default to 1.0 if shcore fails, or calculate if safe
        except:
            return 1.0

def physical_to_logical(px, py):
    """
    Converts physical pixel coordinates (e.g., from ImageGrab) 
    to logical coordinates (e.g., for PyAutoGUI).
    """
    scale = get_dpi_scaling()
    return px / scale, py / scale

def get_virtual_screen_physical_origin():
    """
    Returns (left, top) of the virtual screen in PHYSICAL pixels.
    """
    # win32api metrics are LOGICAL
    logical_left = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    logical_top = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    
    scale = get_dpi_scaling()
    return logical_left * scale, logical_top * scale
