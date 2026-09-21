"""PR form actions with scoped image search and required value readback."""
import ctypes
from ctypes import wintypes
from datetime import datetime
import re

from dateutil.relativedelta import relativedelta
import pyautogui as _input
import pyperclip
import win32gui
import win32process
import roi_helpers

import navigation
import task_control as control
from account_codes import find_index_by_prefix
from image_matcher import Match

pyautogui = control.Input(_input)


def validate(description, account, part):
    if not description.strip():
        raise ValueError('PR 설명을 입력해 주세요.')
    if find_index_by_prefix(account) is None:
        raise ValueError('목록에 있는 계정코드를 선택해 주세요.')
    if part.strip() and len(part.strip()) < 4:
        raise ValueError('Part No는 4자 이상 입력하거나 비워 주세요.')


def focused_value(hwnd):
    """Read standard Windows edit/combo controls without disturbing selection."""
    class GUITHREADINFO(ctypes.Structure):
        _fields_ = [('cbSize', wintypes.DWORD), ('flags', wintypes.DWORD)] + [
            (n, wintypes.HWND) for n in ('hwndActive', 'hwndFocus', 'hwndCapture',
                                        'hwndMenuOwner', 'hwndMoveSize', 'hwndCaret')
        ] + [('rcCaret', wintypes.RECT)]
    control.guard()
    info = GUITHREADINFO()
    info.cbSize = ctypes.sizeof(info)
    tid = win32process.GetWindowThreadProcessId(hwnd)[0]
    if not ctypes.windll.user32.GetGUIThreadInfo(tid, ctypes.byref(info)) or not info.hwndFocus:
        return None
    focus = info.hwndFocus
    if focus != hwnd and not win32gui.IsChild(hwnd, focus):
        return None
    name = win32gui.GetClassName(focus).lower()
    if not any(word in name for word in ('edit', 'combobox', 'sysdatetimepick32')):
        return None
    buf = ctypes.create_unicode_buffer(4096)
    result = ctypes.c_size_t()
    ok = ctypes.windll.user32.SendMessageTimeoutW(
        wintypes.HWND(focus), 0x000D, wintypes.WPARAM(len(buf)),
        ctypes.cast(buf, ctypes.c_void_p), 2, 300, ctypes.byref(result))
    return buf.value if ok else None


def point(hwnd, asset, label, offset=0, timeout=6):
    deadline = control.Clock.monotonic() + timeout
    while control.Clock.monotonic() < deadline:
        control.guard()
        match = navigation.find_item(hwnd, asset, label, deadline=deadline)
        if match:
            scale = getattr(match, 'scale', 1)
            return Match(match[0] + offset * scale, match[1], scale)
        control.sleep(0.15)
    raise RuntimeError(f'{label}: 입력 위치를 찾지 못했습니다. 이후 입력을 중단했습니다.')


def click(hwnd, target):
    if not navigation.mouse(hwnd, target):
        raise RuntimeError('입력 직전 포커스를 확인하지 못했습니다.')


def verify(hwnd, expected, label, kind='text'):
    deadline = control.Clock.monotonic() + 2
    while control.Clock.monotonic() < deadline:
        control.guard()
        value = focused_value(hwnd)
        if value is not None:
            if kind == 'date':
                valid = re.sub(r'\D', '', value) == re.sub(r'\D', '', expected)
            elif kind == 'account':
                valid = value.strip().split('/')[0].strip() == expected
            else:
                valid = value.strip() == expected.strip()
            if valid:
                return True
        control.sleep(0.1)
    # Never report completion based solely on sending keys.
    raise RuntimeError(f'{label}: 입력값을 확인하지 못했습니다. HI-TOPS에서 해당 항목을 확인해 주세요.')


def text_field(hwnd, asset, label, value, offset, kind='text'):
    control.stage(label + ' 입력·확인 중')
    click(hwnd, point(hwnd, asset, label, offset))
    if focused_value(hwnd) is None:
        raise RuntimeError(label + ': 입력 컨트롤을 확인하지 못했습니다.')
    pyautogui.hotkey('ctrl', 'a')
    if kind == 'date':
        pyautogui.press('delete')
    control.guard()
    if kind == 'date':
        pyautogui.write(value, interval=0.01)
    else:
        pyperclip.copy(value)
        pyautogui.hotkey('ctrl', 'v')
    verify(hwnd, value, label, kind)


def click_add(hwnd):
    import pr_add_button
    target, score, scale = pr_add_button.locate(hwnd)
    navigation.log.info('PR Add verified toolbar plus hwnd=%s point=%s score=%.3f scale=%.3f', hwnd, target, score, scale)
    hit = win32gui.WindowFromPoint(tuple(round(v) for v in target))
    if hit != hwnd and not win32gui.IsChild(hwnd, hit):
        raise RuntimeError('다른 창이 + 버튼을 가리고 있어 클릭을 중단했습니다.')
    click(hwnd, target)
    navigation.log.info('PR Add click dispatched; waiting for Detail')


def fill(hwnd, description, unit_price, account, part):
    control.bind_window(hwnd)
    control.stage('새 PR 양식 여는 중')
    click_add(hwnd)
    hwnd = editor_window(hwnd)
    text_field(hwnd, 'description_field.png', 'Description', description, 100)
    date = (datetime.now() + relativedelta(months=1)).strftime('%Y-%m-%d')
    text_field(hwnd, 'need_by_label.png', 'Need By', date, 80, 'date')
    if unit_price:
        control.stage('단가계약 선택·확인 중')
        target = point(hwnd, 'unit_price_label.png', '단가계약', 70)
        click(hwnd, target)
        control.sleep(0.15)
        click(hwnd, (target[0], target[1] + 25 * getattr(target, 'scale', 1)))
        verify(hwnd, 'Y', '단가계약')
    control.stage('계정코드 선택·확인 중')
    click(hwnd, point(hwnd, 'account_code_label.png', 'Account Code', 85))
    control.sleep(0.15)
    pyautogui.press('home')
    for _ in range(find_index_by_prefix(account)):
        pyautogui.press('down')
    pyautogui.press('enter')
    verify(hwnd, account.split('/')[0].strip(), 'Account Code', 'account')
    if part:
        text_field(hwnd, 'part_no_label.png', 'Part No', part, 60)
    control.stage('입력값 확인 완료 · 저장 전 HI-TOPS에서 검토해 주세요')
    return True


def editor_window(previous):
    """Wait for the actual Detail window, never the list's Description filter.

    Add opens an owned WinForms window asynchronously. The list also contains a
    Description field, so matching that image before the new window appears can
    bind the list and immediately trip the foreground guard when Detail opens.
    """
    deadline = control.Clock.monotonic() + 8
    while control.Clock.monotonic() < deadline:
        control.checkpoint()
        foreground = win32gui.GetForegroundWindow()
        for hwnd, title, _ in roi_helpers.application_windows():
            name = title.casefold()
            detail = ('purchase requisition' in name or 'purchase request' in name) and 'detail' in name
            if not detail or hwnd == previous or not win32gui.IsWindowEnabled(hwnd):
                continue
            # The new editor must belong to this list, not a different PR session.
            if win32gui.GetWindow(hwnd, 4) != previous:
                continue
            if foreground == hwnd and win32gui.GetForegroundWindow() == hwnd:
                control.bind_window(hwnd)
                navigation.log.info('PR Detail ready hwnd=%s owner=%s', hwnd, previous)
                return hwnd
        control.sleep(0.15)
    navigation.fail('PR Add clicked but owned foreground Detail was not confirmed')
    raise RuntimeError('새 PR 입력 창을 확인하지 못했습니다. 기존 PR 내용부터 확인해 주세요.')
