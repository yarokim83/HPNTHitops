"""
ROI (Region of Interest) Helpers
Handles window detection and coordinate calculations.
Updated with UIA support for robust window finding.
"""
import win32gui
import win32con
from ctypes import windll
import win32api

import ctypes
from ctypes import wintypes
import ntpath
import win32process

INSTALL_DIR = r"C:\Program Files (x86)\Hyundai-UNI\HITOPSIII"


def menu_caption(menu, index):
    # pywin32 does not expose GetMenuString on every supported build.
    flags = win32con.MF_BYPOSITION
    handle = wintypes.HMENU(menu)
    length = windll.user32.GetMenuStringW(handle, index, None, 0, flags)
    buf = ctypes.create_unicode_buffer(length + 1)
    windll.user32.GetMenuStringW(handle, index, buf, len(buf), flags)
    return buf.value


def _process_path(hwnd):
    handle = None
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        handle = win32api.OpenProcess(0x1000, False, pid)
        size = wintypes.DWORD(32768)
        buf = ctypes.create_unicode_buffer(size.value)
        if windll.kernel32.QueryFullProcessImageNameW(
                wintypes.HANDLE(int(handle)), 0, buf, ctypes.byref(size)):
            return buf.value
    except Exception:
        pass
    finally:
        if handle is not None:
            win32api.CloseHandle(handle)
    return ''


def _is_application_path(path):
    root = ntpath.normcase(ntpath.normpath(INSTALL_DIR)) + '\\'
    return bool(path) and ntpath.normcase(ntpath.normpath(path)).startswith(root)


def application_windows():
    """Visible top-level windows belonging to the configured HI-TOPS installation."""
    found = []
    def collect(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        path = _process_path(hwnd)
        if _is_application_path(path):
            found.append((hwnd, win32gui.GetWindowText(hwnd), path))
    win32gui.EnumWindows(collect, None)
    return found


def _role(title):
    title = title.casefold()
    if 'monitoring' in title or 'm&c' in title:
        return 'mc'
    if 'rcc' in title:
        return 'rcc'
    if 'maintenance' in title and 'repair' in title:
        return 'maintenance'
    if 'hitops' in title or 'hi-tops' in title or 'login' in title:
        return 'main'
    return None


def _window_for_role(role):
    candidates = [(h, title) for h, title, _ in application_windows()
                  if _role(title) == role and win32gui.GetClassName(h) != '#32770']
    # Login may be a native dialog. Include it only for the main/login lookup.
    if role == 'main' and not candidates:
        candidates = [(h, title) for h, title, _ in application_windows()
                      if _role(title) == 'main']
    candidates.sort(key=lambda item: item[0] != win32gui.GetForegroundWindow())
    for hwnd, _ in candidates:
        try:
            return win32gui.GetWindowRect(hwnd), hwnd
        except Exception:
            continue
    return None, None


def get_hitops_window_rect():
    return _window_for_role('main')


def get_mc_window_rect():
    return _window_for_role('mc')


def get_rcc_window_rect():
    return _window_for_role('rcc')


def get_maintenance_window_rect():
    return _window_for_role('maintenance')


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
