import task_control as control
time = control.Clock
"""Verified, bounded navigation for the HI-TOPS Monitoring menus."""
import logging
from logging.handlers import RotatingFileHandler
import os

from PIL import Image, ImageGrab
import pyautogui as _pyautogui
pyautogui = control.Input(_pyautogui)
import win32api
import win32con
import win32gui

import menu_navigator as legacy
import ocr_helpers
import roi_helpers
import image_matcher
import dpi_support

log = logging.getLogger('PRMaker')
_asset_validity = {}


def setup_logging():
    if log.handlers:
        return
    folder = os.path.join(os.getenv('LOCALAPPDATA') or os.path.expanduser('~'), 'PRMaker', 'logs')
    os.makedirs(folder, exist_ok=True)
    handler = RotatingFileHandler(os.path.join(folder, 'navigation.log'),
                                  maxBytes=2_000_000, backupCount=3, encoding='utf-8')
    handler.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    log.addHandler(handler)
    log.setLevel(logging.INFO)


def fail(stage):
    log.error('%s; foreground=%s; windows=%s', stage, win32gui.GetForegroundWindow(),
              roi_helpers.application_windows())
    return False


def activate(hwnd):
    if not hwnd or not win32gui.IsWindow(hwnd):
        return False
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    success = legacy.force_activate_window(hwnd) and win32gui.GetForegroundWindow() == hwnd
    if success:
        control.bind_window(hwnd)
    return success


def find_item(hwnd, asset, label, deadline=None):
    """Search only the visible target window, preserving negative screen origins."""
    control.checkpoint()
    deadline = deadline if deadline is not None else time.monotonic() + 4
    if win32gui.GetForegroundWindow() != hwnd:
        return None
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    # Request only the target window. Coordinates remain absolute on negative-origin monitors.
    crop = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    path = os.path.join(os.path.dirname(__file__), 'assets', asset)
    if path not in _asset_validity:
        try:
            with Image.open(path) as template:
                template.verify()
            _asset_validity[path] = True
        except (OSError, ValueError):
            _asset_validity[path] = False
            log.warning('Missing/invalid menu image: %s; using OCR', asset)
    monitor_scale = dpi_support.monitor_scale(hwnd)
    # Reserve time for OCR when a template cannot match the current appearance.
    image_deadline = min(deadline, time.monotonic() + max(0, (deadline - time.monotonic()) * 0.6))
    box, scale = image_matcher.find(path, crop, confidence=0.8, deadline=image_deadline,
                                  preferred=monitor_scale) if _asset_validity[path] else (None, monitor_scale)
    if box is None:
        scale = monitor_scale
        box = ocr_helpers.find_text_in_image(crop, label, deadline=deadline, source_scale=monitor_scale)
    if box is None:
        return None
    return image_matcher.Match(left + box.left + box.width / 2, top + box.top + box.height / 2, scale)


def mouse(hwnd, point, action='click'):
    if win32gui.GetForegroundWindow() != hwnd:
        return False
    control.bind_window(hwnd)
    legacy.verify_and_execute_mouse(*point, action=action)
    return True


def wait_window(finder, timeout=20):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        # Only handle an identified application error, never send a blind Enter.
        if error_dialog():
            return None
        _, hwnd = finder()
        if hwnd:
            return hwnd
        time.sleep(0.3)
    return None


def error_dialog():
    """Detect scoped error dialogs and leave their contents visible to the user."""
    keywords = ('error', 'authority', 'access denied', 'warning', '오류', '권한', '경고')
    for hwnd, title, _ in roi_helpers.application_windows():
        if win32gui.GetClassName(hwnd) == '#32770' and any(k in title.lower() for k in keywords):
            log.error('Application error dialog hwnd=%s title=%r', hwnd, title)
            return hwnd
    return None


def open_monitoring(target, finder):
    control.stage(target + ' 메뉴 여는 중')
    if not legacy.ensure_app_ready():
        fail('HI-TOPS initialization failed')
        return None
    hwnd = legacy.login_manager.main_window()
    if not hwnd:
        fail('HI-TOPS main disappeared after login')
        return None
    for attempt in range(3):
        log.info('Monitoring -> %s attempt=%d hwnd=%s', target, attempt + 1, hwnd)
        if error_dialog() or not activate(hwnd):
            fail('Cannot activate HI-TOPS or error dialog present')
            return None
        pyautogui.press('esc')
        root = find_item(hwnd, 'monitoring_menu.png', 'Monitoring')
        if not root:
            time.sleep(0.5)
            continue
        # Leave the previous hover target before re-entering it.
        rect = win32gui.GetWindowRect(hwnd)
        pyautogui.moveTo(rect[0] + 30, rect[1] + 50, duration=0.1)
        if not mouse(hwnd, root, 'hover'):
            return None
        deadline = time.monotonic() + 5
        while time.monotonic() < deadline:
            point = find_item(hwnd, 'mc_menu_item.png' if target == 'M&C' else 'rcc_icon.png', target, deadline=deadline)
            if point:
                if not mouse(hwnd, point):
                    return None
                result = wait_window(finder)
                if result:
                    return result
                # A click was sent: do not launch duplicate windows on a slow startup.
                fail('%s window did not appear after click' % target)
                return None
            time.sleep(0.3)
    fail('Monitoring submenu not found: ' + target)
    return None


def schedule_visible(hwnd):
    """Require an actual top-level/MDI window caption, not the open menu text."""
    handles = [hwnd]
    win32gui.EnumChildWindows(hwnd, lambda h, out: out.append(h), handles)
    handles.extend(h for h, _, _ in roi_helpers.application_windows())
    for h in handles:
        if (win32gui.IsWindowVisible(h)
                and win32gui.GetClassName(h) != '#32768'
                and 'berthing schedule' in win32gui.GetWindowText(h).lower()):
            return True
    return False


def native_schedule(hwnd):
    """Use the actual enabled command ID where the application exposes an HMENU."""
    def search(menu):
        for index in range(win32gui.GetMenuItemCount(menu)):
            caption = roi_helpers.menu_caption(menu, index)
            caption = caption.replace('&', '').split('\t')[0].strip().rstrip('.').casefold()
            state = win32gui.GetMenuState(menu, index, win32con.MF_BYPOSITION)
            if state & (win32con.MF_DISABLED | win32con.MF_GRAYED):
                continue
            sub = win32gui.GetSubMenu(menu, index)
            if sub:
                command = search(sub)
                if command is not None:
                    return command
            elif caption == 'berthing schedule':
                command = win32gui.GetMenuItemID(menu, index)
                if command not in (-1, 0xFFFFFFFF):
                    return command
        return None
    try:
        menu = win32gui.GetMenu(hwnd)
        command = search(menu) if menu else None
    except Exception:
        log.debug('Native menu unavailable', exc_info=True)
        return False
    if command is None or win32gui.GetForegroundWindow() != hwnd:
        return False
    control.guard()
    win32gui.PostMessage(hwnd, win32con.WM_COMMAND, command, 0)
    log.info('Berthing Schedule native command=%s', command)
    return True


def wait_schedule(hwnd):
    deadline = time.monotonic() + 15
    while time.monotonic() < deadline:
        if error_dialog():
            return fail('Berthing Schedule blocked by error dialog')
        if schedule_visible(hwnd):
            log.info('Berthing Schedule window confirmed')
            return True
        time.sleep(0.3)
    return fail('Selection sent, but Berthing Schedule window not confirmed')


def run_mc():
    setup_logging()
    control.stage('M&C 창 확인 중')
    _, hwnd = roi_helpers.get_mc_window_rect()
    hwnd = hwnd or open_monitoring('M&C', roi_helpers.get_mc_window_rect)
    if not hwnd or error_dialog() or not activate(hwnd):
        return fail('Cannot activate M&C')
    control.stage('Berthing Schedule 여는 중')
    if schedule_visible(hwnd):
        return True
    # Vessel menu order verified against the user's HI-TOPS menu (18 items).
    # Home resets selection even if a previous run left this menu open.
    # Each individual key is focus/cancellation guarded by control.Input.
    pyautogui.press('esc')
    pyautogui.hotkey('alt', 'v')
    time.sleep(0.4)
    pyautogui.press('home')
    for _ in range(9):
        pyautogui.press('down')
        time.sleep(0.08)
    if error_dialog():
        return fail('Berthing Schedule blocked before keyboard selection')
    pyautogui.press('enter')
    log.info('Berthing Schedule keyboard selection: Alt+V, Home, Down x9, Enter')
    # Do not repeat Enter after a slow response: confirm the opened window.
    return wait_schedule(hwnd)


def run_rcc():
    setup_logging()
    control.stage('RCC 창 확인 중')
    _, hwnd = roi_helpers.get_rcc_window_rect()
    hwnd = hwnd or open_monitoring('RCC', roi_helpers.get_rcc_window_rect)
    if not hwnd or error_dialog() or not activate(hwnd):
        return fail('RCC window not confirmed/activated')
    log.info('RCC window confirmed hwnd=%s', hwnd)
    return True
