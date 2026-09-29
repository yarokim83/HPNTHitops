"""Colour-aware plus matching restricted to the PR client toolbar."""
import logging
from pathlib import Path
import cv2
import numpy as np
from PIL import Image, ImageGrab
import win32gui
import task_control as control
from dpi_support import monitor_scale, candidate_scales


def find_plus(image, preferred=1):
    source = np.asarray(image.convert('RGB'))
    with Image.open(Path(__file__).parent / 'assets' / 'add_btn.png') as asset:
        original = asset.convert('RGB')
    best = None
    for scale in candidate_scales(preferred):
        control.checkpoint()
        width, height = (round(size * scale) for size in original.size)
        if min(width, height) < 12 or width > image.width or height > image.height:
            continue
        template = np.asarray(original.resize((width, height), Image.Resampling.LANCZOS))
        scores = cv2.matchTemplate(source, template, cv2.TM_CCOEFF_NORMED)
        _, score, _, (x, y) = cv2.minMaxLoc(scores)
        if score < 0.88:
            continue
        patch = source[y:y+height, x:x+width].astype(np.int16)
        green = (patch[:,:,1] > patch[:,:,0] + 20) & (patch[:,:,1] > patch[:,:,2] + 10)
        if green.mean() < 0.06:
            continue
        if best is None or score > best[2]:
            best = (x + width / 2, y + height / 2, score, scale)
    return best


def locate(hwnd, timeout=6):
    """Wait for painting/layout to settle; never send input while searching."""
    log = logging.getLogger('PRMaker')
    deadline = control.Clock.monotonic() + timeout
    previous = None
    attempts = 0
    box, scale = None, None
    while control.Clock.monotonic() < deadline:
        control.guard()
        if not win32gui.IsWindow(hwnd):
            raise RuntimeError('PR 목록 창이 닫혔습니다. 클릭하지 않았습니다.')
        if not win32gui.IsWindowEnabled(hwnd) or win32gui.IsIconic(hwnd):
            previous = None
            control.sleep(0.2)
            continue
        scale = monitor_scale(hwnd)
        left, top = win32gui.ClientToScreen(hwnd, (0, 0))
        _, _, width, height = win32gui.GetClientRect(hwnd)
        # Allow padding/two toolbar rows at different DPI, without scanning tables.
        box = (left, top, left + min(width, round(300 * scale)),
               top + min(height, round(64 * scale)))
        if width <= 0 or height <= 0:
            previous = None
            control.sleep(0.2)
            continue
        attempts += 1
        match = find_plus(ImageGrab.grab(bbox=box, all_screens=True), scale)
        control.guard()
        if match is not None:
            x, y, score, matched_scale = match
            target = (left + x, top + y)
            # Repaint or movement may change coordinates: require two frames at
            # the same absolute position before handing a single click to caller.
            if previous and previous[0] == box and all(abs(a-b) <= 2 for a,b in zip(previous[1], target)):
                log.info('PR Add stable detection hwnd=%s attempts=%s bbox=%s dpi=%.3f score=%.3f',
                         hwnd, attempts, box, scale, score)
                return target, score, matched_scale
            previous = (box, target)
        else:
            previous = None
        control.sleep(0.2)
    log.warning('PR Add detection timed out hwnd=%s attempts=%s bbox=%s dpi=%s foreground=%s',
                hwnd, attempts, box, scale, win32gui.GetForegroundWindow())
    raise RuntimeError('상단 초록색 + 버튼을 6초 동안 확인하지 못했습니다. 클릭하지 않았습니다.')
