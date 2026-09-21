"""Colour-aware plus matching restricted to the PR client toolbar."""
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


def locate(hwnd):
    control.guard()
    if not win32gui.IsWindowEnabled(hwnd) or win32gui.IsIconic(hwnd):
        raise RuntimeError('PR 목록 창이 입력 가능한 상태가 아닙니다.')
    scale = monitor_scale(hwnd)
    left, top = win32gui.ClientToScreen(hwnd, (0, 0))
    _, _, width, height = win32gui.GetClientRect(hwnd)
    # Only the first toolbar row, excluding filters and PR/Part tables.
    box = (left, top, left + min(width, round(300 * scale)),
           top + min(height, round(32 * scale)))
    match = find_plus(ImageGrab.grab(bbox=box, all_screens=True), scale)
    if match is None:
        raise RuntimeError('상단 도구모음의 초록색 + 버튼을 확인하지 못했습니다. 클릭하지 않았습니다.')
    x, y, score, matched_scale = match
    return (left + x, top + y), score, matched_scale
