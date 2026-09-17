"""Consistent physical coordinates and scale hints for mixed-DPI desktops."""
import ctypes
from ctypes import wintypes


def enable_per_monitor():
    # Must run before importing Tk/PyAutoGUI, which otherwise choose process DPI mode.
    try:
        return bool(ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)))
    except (AttributeError, OSError):
        try:
            return ctypes.windll.shcore.SetProcessDpiAwareness(2) == 0
        except (AttributeError, OSError):
            return False


def monitor_scale(hwnd):
    try:
        import win32api
        import win32con
        monitor = win32api.MonitorFromWindow(hwnd, win32con.MONITOR_DEFAULTTONEAREST)
        scale = ctypes.c_int()
        if ctypes.windll.shcore.GetScaleFactorForMonitor(
                wintypes.HANDLE(int(monitor)), ctypes.byref(scale)) == 0 and scale.value > 0:
            return scale.value / 100
    except (AttributeError, OSError, TypeError):
        pass
    try:
        dpi = ctypes.windll.user32.GetDpiForWindow(wintypes.HWND(hwnd))
        return dpi / 96 if dpi else 1
    except (AttributeError, OSError):
        return 1


def candidate_scales(preferred=1, previous=None):
    """Support assets captured on a different monitor from the target application."""
    references = (1, 1.25, 1.5, 1.75, 2, 2.25, 2.5, 3)
    candidates = [previous, preferred, 1, preferred / 1.25]
    candidates += [preferred / reference for reference in references]
    candidates += list(references) + [0.8, 2 / 3, 0.5, 4 / 7]
    result = []
    for value in candidates:
        if value is None or not 0.33 <= value <= 3:
            continue
        if not any(abs(value - prior) < 0.01 for prior in result):
            result.append(value)
    return result
