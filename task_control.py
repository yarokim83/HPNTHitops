"""Per-worker cancellation, progress and foreground input ownership."""
import contextlib
import threading
import time as _time
import logging


class TaskStopped(BaseException):
    """Bypass legacy broad exception handlers: no further input may be sent."""


_local = threading.local()


@contextlib.contextmanager
def session(cancel, progress):
    _local.cancel, _local.progress, _local.hwnd = cancel, progress, None
    try:
        yield
    finally:
        _local.__dict__.clear()


def checkpoint():
    event = getattr(_local, 'cancel', None)
    if event and event.is_set():
        raise TaskStopped('사용자가 중지했습니다. 이미 입력된 내용은 HI-TOPS에서 확인해 주세요.')


def stage(text):
    checkpoint()
    logging.getLogger('PRMaker').info('stage=%s', text)
    callback = getattr(_local, 'progress', None)
    if callback:
        callback(text)


def bind_window(hwnd):
    checkpoint()
    _local.hwnd = hwnd


def guard():
    checkpoint()
    hwnd = getattr(_local, 'hwnd', None)
    if hwnd:
        import win32gui
        foreground = win32gui.GetForegroundWindow()
        if foreground != hwnd:
            logging.getLogger('PRMaker').warning('Input focus changed expected=%s actual=%s', hwnd, foreground)
            raise TaskStopped('대상 창의 포커스가 바뀌어 중지했습니다. 이미 입력된 내용을 확인해 주세요.')


def sleep(seconds):
    checkpoint()
    event = getattr(_local, 'cancel', None)
    if event:
        if event.wait(seconds):
            checkpoint()
    else:
        _time.sleep(seconds)


class Clock:
    sleep = staticmethod(sleep)
    monotonic = staticmethod(_time.monotonic)
    time = staticmethod(_time.time)


class Input:
    ACTIONS = {'click', 'doubleClick', 'moveTo', 'moveRel', 'hotkey', 'press',
               'write', 'typewrite', 'keyDown', 'keyUp', 'scroll', 'dragTo'}

    def __init__(self, api):
        object.__setattr__(self, '_api', api)

    def __getattr__(self, name):
        value = getattr(self._api, name)
        if name not in self.ACTIONS:
            return value
        def call(*args, **kwargs):
            guard()
            return value(*args, **kwargs)
        return call

    def __setattr__(self, name, value):
        setattr(self._api, name, value)
