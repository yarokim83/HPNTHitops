"""Headless regression tests: no ERP, mouse, keyboard or credentials are accessed.

Run: python -m unittest discover -s tests -p '*_regression.py' -v
"""
import ast
import itertools
import ntpath
import os
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

ROOT = Path(__file__).resolve().parents[1]


def functions(file, *names, **namespace):
    tree = ast.parse((ROOT / file).read_text(encoding='utf-8-sig'))
    tree.body = [node for node in tree.body
                 if isinstance(node, (ast.FunctionDef, ast.ClassDef)) and node.name in names]
    exec(compile(tree, file, 'exec'), namespace)
    return SimpleNamespace(**namespace)


def nav(*names, **overrides):
    clock = itertools.count()
    base = dict(time=SimpleNamespace(monotonic=lambda: next(clock) / 2, sleep=Mock()),
                log=Mock(), pyautogui=Mock(), setup_logging=Mock(), fail=Mock(return_value=False), control=Mock(), image_matcher=SimpleNamespace(Match=lambda x,y,scale=1: (x,y)), dpi_support=SimpleNamespace(monitor_scale=lambda h: 1))
    base.update(overrides)
    return functions('navigation.py', *names, **base)


class WindowTests(unittest.TestCase):
    def test_mc_is_not_excluded_when_it_contains_hitops(self):
        api = Mock()
        api.GetClassName.return_value = 'Application'
        api.GetWindowRect.return_value = (0, 0, 1920, 1080)
        api.GetForegroundWindow.return_value = 1
        mod = functions('roi_helpers.py', '_role', '_window_for_role',
                        'get_hitops_window_rect', 'get_mc_window_rect', win32gui=api,
                        application_windows=lambda: [(1, 'HI-TOPS Monitoring & Control', ''),
                                                     (2, 'HI-TOPS Main', '')])
        self.assertEqual(mod.get_hitops_window_rect()[1], 2)
        self.assertEqual(mod.get_mc_window_rect()[1], 1)

    def test_unrelated_monitor_process_is_excluded(self):
        api = Mock()
        api.IsWindowVisible.return_value = True
        api.GetWindowText.side_effect = lambda h: {1: 'Resource Monitor', 2: 'M&C'}[h]
        api.EnumWindows.side_effect = lambda cb, ctx: [cb(h, ctx) for h in (1, 2)]
        mod = functions('roi_helpers.py', '_is_application_path', 'application_windows',
                        ntpath=ntpath, INSTALL_DIR=r'C:\HI-TOPS', win32gui=api,
                        _process_path=lambda h: {1: r'C:\Windows\monitor.exe',
                                                2: r'C:\HI-TOPS\mc.exe'}[h])
        self.assertEqual([w[0] for w in mod.application_windows()], [2])
        self.assertFalse(mod._is_application_path(r'C:\HI-TOPS-fake\mc.exe'))

    def test_title_only_does_not_imply_login_based_on_size(self):
        api = Mock()
        api.GetWindowText.return_value = 'HI-TOPS Main'
        api.EnumChildWindows.side_effect = lambda h, cb, out: None
        mod = functions('login_manager.py', 'is_login', win32gui=api, win32con=Mock())
        self.assertFalse(mod.is_login(10))


class NavigationTests(unittest.TestCase):
    def test_crop_restores_negative_monitor_coordinates(self):
        shot = Mock(width=3840, height=1080)
        api = SimpleNamespace(GetForegroundWindow=lambda: 1,
                              GetWindowRect=lambda h: (-1600, 0, -400, 900))
        path = os.path.join(str(ROOT), 'assets', 'test.png')
        mod = nav('find_item', win32gui=api,
                  win32con=SimpleNamespace(SM_XVIRTUALSCREEN=76, SM_YVIRTUALSCREEN=77),
                  win32api=SimpleNamespace(GetSystemMetrics=lambda key: -1920 if key == 76 else 0),
                  ImageGrab=SimpleNamespace(grab=lambda **k: shot),
                  image_matcher=SimpleNamespace(find=lambda *a, **k: (
                                         SimpleNamespace(left=100, top=20, width=40, height=20), 1),
                                               Match=lambda x,y,scale: (x,y)),
                  os=os, __file__=str(ROOT / 'navigation.py'), _asset_validity={path: True})
        self.assertEqual(mod.find_item(1, 'test.png', 'test'), (-1480, 30))
        shot.crop.assert_not_called()

    def test_activation_retry_must_recheck_foreground(self):
        api = Mock()
        api.IsWindow.return_value = True
        api.IsIconic.return_value = False
        api.GetForegroundWindow.return_value = 99
        fake_ctypes = Mock()
        fake_ctypes.windll.kernel32.GetCurrentThreadId.return_value = 10
        mod = functions('menu_navigator.py', 'force_activate_window',
                        win32gui=api, win32con=Mock(), ctypes=fake_ctypes, control=Mock(),
                        win32process=SimpleNamespace(GetWindowThreadProcessId=lambda h: (20, 30)),
                        time=SimpleNamespace(sleep=Mock()))
        self.assertFalse(mod.force_activate_window(1))
        fake_ctypes.windll.user32.AttachThreadInput.assert_any_call(10, 20, False)

    def test_no_keyboard_input_after_activation_failure(self):
        mod = nav('run_mc', roi_helpers=SimpleNamespace(get_mc_window_rect=lambda: (None, 1)),
                  error_dialog=lambda: None, activate=lambda h: False)
        self.assertFalse(mod.run_mc())
        mod.pyautogui.hotkey.assert_not_called()
        mod.pyautogui.press.assert_not_called()

    def test_missing_menu_retries_without_down_or_enter(self):
        mod = nav('run_mc', roi_helpers=SimpleNamespace(get_mc_window_rect=lambda: (None, 1)),
                  error_dialog=lambda: None, activate=lambda h: True,
                  native_schedule=lambda h: False, find_item=Mock(return_value=None))
        self.assertFalse(mod.run_mc())
        self.assertEqual(mod.pyautogui.hotkey.call_count, 3)
        self.assertEqual([c.args for c in mod.pyautogui.press.call_args_list], [('esc',)] * 3)

    def test_success_requires_schedule_confirmation(self):
        for confirmed in (False, True):
            mod = nav('run_mc', roi_helpers=SimpleNamespace(get_mc_window_rect=lambda: (None, 1)),
                      error_dialog=lambda: None, activate=lambda h: True,
                      native_schedule=lambda h: False, find_item=lambda *args, **kwargs: (40, 50),
                      mouse=lambda *args: True, wait_schedule=lambda h: confirmed)
            self.assertEqual(mod.run_mc(), confirmed)

    def test_submenu_reenters_hover_after_first_timeout(self):
        calls = []
        def find(hwnd, asset, label, **kwargs):
            if label == 'Monitoring':
                calls.append(label)
                return (10, 20)
            return (30, 40) if len(calls) == 2 else None
        mod = nav('open_monitoring', legacy=SimpleNamespace(ensure_app_ready=lambda: True),
                  roi_helpers=SimpleNamespace(get_hitops_window_rect=lambda: (None, 1)),
                  win32gui=SimpleNamespace(GetWindowRect=lambda h: (0, 0, 800, 600)),
                  error_dialog=lambda: None, activate=lambda h: True, find_item=find,
                  mouse=Mock(return_value=True), wait_window=lambda finder: 2)
        self.assertEqual(mod.open_monitoring('M&C', Mock()), 2)
        self.assertEqual(len(calls), 2)

    def test_click_timeout_does_not_launch_duplicate_windows(self):
        mouse = Mock(return_value=True)
        mod = nav('open_monitoring', legacy=SimpleNamespace(ensure_app_ready=lambda: True),
                  roi_helpers=SimpleNamespace(get_hitops_window_rect=lambda: (None, 1)),
                  win32gui=SimpleNamespace(GetWindowRect=lambda h: (0, 0, 800, 600)),
                  error_dialog=lambda: None, activate=lambda h: True,
                  find_item=lambda *a, **k: (30, 40), mouse=mouse, wait_window=lambda finder: None)
        self.assertIsNone(mod.open_monitoring('M&C', Mock()))
        self.assertEqual(mouse.call_count, 2)  # one hover, one click

    def test_native_menu_uses_name_not_index_and_respects_disabled(self):
        api = Mock()
        api.GetMenu.return_value = 50
        api.GetMenuItemCount.return_value = 1
        api.GetMenuString.return_value = '&Berthing Schedule\tCtrl+B'
        api.GetSubMenu.return_value = 0
        api.GetMenuItemID.return_value = 321
        api.GetForegroundWindow.return_value = 1
        constants = SimpleNamespace(MF_BYPOSITION=1024, MF_DISABLED=2, MF_GRAYED=1, WM_COMMAND=273)
        mod = nav('native_schedule', win32gui=api, win32con=constants,
                  roi_helpers=SimpleNamespace(menu_caption=api.GetMenuString))
        api.GetMenuState.return_value = 0
        self.assertTrue(mod.native_schedule(1))
        api.PostMessage.assert_called_once_with(1, 273, 321, 0)
        api.PostMessage.reset_mock()
        api.GetMenuState.return_value = 2
        self.assertFalse(mod.native_schedule(1))
        api.PostMessage.assert_not_called()

    def test_dpi_fallback_when_library_returns_none(self):
        fallback = Mock(return_value='box')
        api = SimpleNamespace(locate=Mock(return_value=None), ImageNotFoundException=RuntimeError)
        mod = functions('menu_navigator.py', 'safe_locate', pyautogui=api,
                        locate_with_scaling=fallback)
        self.assertEqual(mod.safe_locate('asset', 'screen'), 'box')
        self.assertIn(1.5, fallback.call_args.kwargs['scales'])

    def test_mouse_position_failure_aborts_click(self):
        api = Mock()
        api.position.return_value = (1000, 1000)
        mod = functions('menu_navigator.py', 'verify_and_execute_mouse',
                        pyautogui=api, win32api=Mock(), time=SimpleNamespace(sleep=Mock()), control=Mock())
        with self.assertRaises(RuntimeError):
            mod.verify_and_execute_mouse(10, 10)
        api.click.assert_not_called()


class OcrTests(unittest.TestCase):
    def test_phrase_tokens_on_same_line_are_combined(self):
        data = dict(text=['Berthing', 'Schedule'], conf=['92.5', '94'],
                    block_num=[1, 1], par_num=[1, 1], line_num=[1, 1],
                    left=[10, 100], top=[20, 20], width=[80, 90], height=[20, 20])
        fake = SimpleNamespace(image_to_data=lambda *a, **k: data,
                               Output=SimpleNamespace(DICT='dict'))
        mod = functions('ocr_helpers.py', '_OcrBox', '_scan_for_matches',
                        _preprocess_extreme=lambda *a, **k: ('image', 2), control=Mock())
        with patch.dict(sys.modules, pytesseract=fake):
            boxes = mod._scan_for_matches('shot', 'Berthing Schedule', 'full', False)
            self.assertEqual(len(boxes), 1)
            self.assertEqual((boxes[0].left, boxes[0].width), (5, 90))
            data['line_num'] = [1, 2]
            self.assertEqual(mod._scan_for_matches('shot', 'Berthing Schedule', 'full', False), [])


if __name__ == '__main__':
    unittest.main()
