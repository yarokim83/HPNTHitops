import task_control as control
time = control.Clock
import pyautogui as _pyautogui
pyautogui = control.Input(_pyautogui)
import pyperclip
import win32gui
import win32con
import roi_helpers

# Global configuration
pyautogui.FAILSAFE = True

def window_enum_handler(hwnd, resultList):
    if win32gui.IsWindowVisible(hwnd):
        resultList.append((hwnd, win32gui.GetWindowText(hwnd)))

def get_app_window(partial_title_list):
    for hwnd, title, _ in roi_helpers.application_windows():
        if any(part.casefold() in title.casefold() for part in partial_title_list):
            return hwnd, title
    return None, None


def is_login(hwnd):
    if 'login' in win32gui.GetWindowText(hwnd).lower():
        return True
    fields = []
    def inspect(child, out):
        # Standard password edits, including WindowsForms edit class names.
        if ('edit' in win32gui.GetClassName(child).lower()
                and win32gui.IsWindowVisible(child)
                and win32gui.GetWindowLong(child, win32con.GWL_STYLE) & 0x20):
            out.append(child)
    win32gui.EnumChildWindows(hwnd, inspect, fields)
    return bool(fields)


def main_window():
    """Inspect every candidate; a lingering Login window must not hide main."""
    for hwnd, title, _ in roi_helpers.application_windows():
        name = title.casefold()
        if ('hitops' not in name and 'hi-tops' not in name) or 'login' in name:
            continue
        try:
            if (win32gui.GetClassName(hwnd) != '#32770'
                    and win32gui.IsWindowEnabled(hwnd) and not is_login(hwnd)):
                return hwnd
        except Exception:
            continue  # A login window may disappear while enumerating.
    return None


def wait_main(timeout=30):
    import logging
    log = logging.getLogger('PRMaker')
    control.stage('로그인 후 HI-TOPS 메인 창 기다리는 중')
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        control.checkpoint()
        hwnd = main_window()
        if hwnd:
            log.info('Login main window ready hwnd=%s', hwnd)
            return True
        time.sleep(0.3)
    log.error('Login main wait timed out; foreground=%s windows=%s',
              win32gui.GetForegroundWindow(), roi_helpers.application_windows())
    return False


def perform_login(password):
    import logging
    from menu_navigator import force_activate_window
    log = logging.getLogger('PRMaker')
    deadline = time.monotonic() + 30
    hwnd = None
    while time.monotonic() < deadline:
        control.checkpoint()
        if main_window():
            log.info('Login: existing HI-TOPS main confirmed')
            return True
        for candidate, _, _ in roi_helpers.application_windows():
            try:
                if is_login(candidate):
                    hwnd = candidate
                    break
            except Exception:
                continue
        if hwnd:
            break
        time.sleep(0.3)
    if not hwnd:
        log.error('Login: no login or main window appeared')
        return False

    # Retry activation without resubmitting credentials. The user/app may already
    # be completing login while the old login handle disappears.
    control.stage('로그인 창 활성화 중')
    for attempt in range(3):
        if main_window():
            return True
        if not win32gui.IsWindow(hwnd):
            return wait_main()
        if force_activate_window(hwnd):
            break
        log.warning('Login activation pending attempt=%s hwnd=%s', attempt + 1, hwnd)
        time.sleep(0.3)
    else:
        log.warning('Login activation unavailable; waiting for main without sending credentials')
        return wait_main()

    if main_window():
        return True
    try:
        if not win32gui.IsWindow(hwnd) or not is_login(hwnd):
            return wait_main()
        if win32gui.GetForegroundWindow() != hwnd:
            log.warning('Login focus changed before input; waiting for main')
            return wait_main()
        control.bind_window(hwnd)
        rect = win32gui.GetWindowRect(hwnd)
        pyautogui.click((rect[0] + rect[2]) // 2, (rect[1] + rect[3]) // 2)
        time.sleep(0.1)
        if win32gui.GetForegroundWindow() != hwnd:
            log.warning('Login focus changed after field click; waiting for main')
            return wait_main()
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.write(password, interval=0.01)
        pyautogui.press('enter')
        log.info('Login submission sent once; awaiting main window')
    except Exception as exc:
        # Never log credential text or exception messages that might contain it.
        log.warning('Login input failed type=%s; awaiting main without resubmission', type(exc).__name__)
    return wait_main()
