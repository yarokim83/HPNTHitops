"""Bounded image matching with cached decoded and resized templates."""
from functools import lru_cache
import os
import time
from PIL import Image
import pyautogui
import task_control as control
from dpi_support import candidate_scales

_last_scale = {}


@lru_cache(maxsize=128)
def template(path, stamp, scale):
    with Image.open(path) as image:
        image = image.convert('RGB')
        if scale != 1:
            image = image.resize((max(1, round(image.width * scale)),
                                  max(1, round(image.height * scale))), Image.Resampling.LANCZOS)
        return image


class Match(tuple):
    def __new__(cls, x, y, scale=1):
        obj = super().__new__(cls, (x, y))
        obj.scale = scale
        return obj


def find(path, image, confidence=0.8, deadline=None, preferred=1):
    stamp = os.stat(path).st_mtime_ns
    key = (path, stamp, round(preferred, 2))
    for scale in candidate_scales(preferred, _last_scale.get(key)):
        control.checkpoint()
        if deadline is not None and time.monotonic() >= deadline:
            break
        needle = template(path, stamp, scale)
        if needle.width > image.width or needle.height > image.height:
            continue
        try:
            box = pyautogui.locate(needle, image, confidence=confidence)
            if box is not None:
                _last_scale[key] = scale
                return box, scale
        except pyautogui.ImageNotFoundException:
            pass
    return None, 1
