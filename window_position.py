"""Place Tk top-levels in absolute Windows coordinates, including negative origins."""
import win32api
import win32con
import win32gui


def clamp(x, y, width, height, work):
    left, top, right, bottom = work
    return max(left, min(x, right - width)), max(top, min(y, bottom - height))


def place(window, x, y):
    window.update_idletasks()
    monitor = win32api.MonitorFromPoint((int(x), int(y)), win32con.MONITOR_DEFAULTTONEAREST)
    work = win32api.GetMonitorInfo(monitor)['Work']
    hwnd = win32gui.GetAncestor(window.winfo_id(), win32con.GA_ROOT)
    rect = win32gui.GetWindowRect(hwnd)
    width, height = rect[2] - rect[0], rect[3] - rect[1]
    x, y = clamp(int(x), int(y), width, height, work)
    win32gui.SetWindowPos(hwnd, 0, x, y, 0, 0,
                          win32con.SWP_NOSIZE | win32con.SWP_NOZORDER | win32con.SWP_NOACTIVATE)
