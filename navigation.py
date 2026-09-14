"""Verified, bounded navigation for the HI-TOPS Monitoring menus."""
import logging
from logging.handlers import RotatingFileHandler
import os
import time

from PIL import Image, ImageGrab
import pyautogui
import win32api
import win32con
import win32gui

import menu_navigator as legacy
import ocr_helpers
import roi_helpers

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
    return legacy.force_activate_window(hwnd) and win32gui.GetForegroundWindow() == hwnd


def find_item(hwnd, asset, label):
    """Search only the visible target window, preserving negative screen origins."""
    if win32gui.GetForegroundWindow() != hwnd:
        return None
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    vx = win32api.GetSystemMetrics(win32con.SM_XVIRTUALSCREEN)
    vy = win32api.GetSystemMetrics(win32con.SM_YVIRTUALSCREEN)
    shot = ImageGrab.grab(all_screens=True)
    left, top = max(left, vx), max(top, vy)
    right, bottom = min(right, vx + shot.width), min(bottom, vy + shot.height)
    if right <= left or bottom <= top:
        return None
    crop = shot.crop((left - vx, top - vy, right - vx, bottom - vy))
    path = os.path.join(os.path.dirname(__file__), 'assets', asset)
    if path not in _asset_validity:
        try:
            with Image.open(path) as template:
                template.verify()
            _asset_validity[path] = True
        except (OSError, ValueError):
            _asset_validity[path] = False
            log.warning('Missing/invalid menu image: %s; using OCR', asset)
    box = legacy.safe_locate(path, crop, confidence=0.8) if _asset_validity[path] else None
    if box is None:
        box = ocr_helpers.find_text_in_image(crop, label)
    if box is None:
        return None
    return left + box.left + box.width / 2, top + box.top + box.height / 2


def mouse(hwnd, point, action='click'):
    if win32gui.GetForegroundWindow() != hwnd:
        return False
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
    if not legacy.ensure_app_ready():
        fail('HI-TOPS initialization failed')
        return None
    _, hwnd = roi_helpers.get_hitops_window_rect()
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
            point = find_item(hwnd, 'mc_menu_item.png' if target == 'M&C' else 'rcc_icon.png', target)
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
    _, hwnd = roi_helpers.get_mc_window_rect()
    hwnd = hwnd or open_monitoring('M&C', roi_helpers.get_mc_window_rect)
    if not hwnd or error_dialog() or not activate(hwnd):
        return fail('Cannot activate M&C')
    if native_schedule(hwnd):
        return wait_schedule(hwnd)
    for attempt in range(3):
        if not activate(hwnd):
            return fail('M&C lost focus')
        pyautogui.press('esc')
        pyautogui.hotkey('alt', 'v')
        deadline = time.monotonic() + 4
        while time.monotonic() < deadline:
            point = find_item(hwnd, 'berthing_schedule.png', 'Berthing Schedule')
            if point:
                if not mouse(hwnd, point):
                    return fail('M&C lost focus before selection')
                return wait_schedule(hwnd)
            time.sleep(0.3)
        log.warning('Vessel menu/item not found; retry=%d', attempt + 1)
    return fail('Berthing Schedule menu item not found')


def run_rcc():
    setup_logging()
    _, hwnd = roi_helpers.get_rcc_window_rect()
    hwnd = hwnd or open_monitoring('RCC', roi_helpers.get_rcc_window_rect)
    if not hwnd or error_dialog() or not activate(hwnd):
        return fail('RCC window not confirmed/activated')
    log.info('RCC window confirmed hwnd=%s', hwnd)
    return True
