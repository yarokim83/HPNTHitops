import os
import sys
import threading
import unittest
import itertools
from types import SimpleNamespace
from unittest.mock import Mock, patch

sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
import task_control as control
from navigation_regression import functions, nav


class WorkflowTests(unittest.TestCase):
    def test_existing_maintenance_skips_main_and_tile(self):
        navigation = Mock()
        navigation.activate.return_value = True
        navigation.mouse.return_value = True
        navigation.find_item.return_value = (10, 20)
        mod = functions('menu_navigator.py', 'click_pr_menu', control=Mock(),
                        roi_helpers=SimpleNamespace(get_maintenance_window_rect=lambda: (None, 7)),
                        win32gui=Mock(), ensure_app_ready=Mock(),
                        time=SimpleNamespace(monotonic=lambda: 0, sleep=Mock()))
        with patch.dict(sys.modules, navigation=navigation):
            self.assertTrue(mod.click_pr_menu())
        mod.ensure_app_ready.assert_not_called()
        navigation.activate.assert_called_once_with(7)
        self.assertEqual([c.args[2] for c in navigation.find_item.call_args_list],
                         ['Inventory', 'Purchase Request'])

    def test_existing_maintenance_activation_failure_sends_no_input(self):
        navigation = Mock()
        navigation.activate.return_value = False
        navigation.fail.return_value = False
        mod = functions('menu_navigator.py', 'click_pr_menu', control=Mock(),
                        roi_helpers=SimpleNamespace(get_maintenance_window_rect=lambda: (None, 7)),
                        win32gui=Mock(), ensure_app_ready=Mock())
        with patch.dict(sys.modules, navigation=navigation):
            self.assertFalse(mod.click_pr_menu())
        navigation.find_item.assert_not_called()
        navigation.mouse.assert_not_called()
        mod.ensure_app_ready.assert_not_called()

    def test_main_activation_required_before_tile_search(self):
        navigation = Mock()
        navigation.activate.return_value = False
        navigation.fail.return_value = False
        mod = functions('menu_navigator.py', 'click_pr_menu', control=Mock(),
                        roi_helpers=SimpleNamespace(get_maintenance_window_rect=lambda: (None, None),
                                                    get_hitops_window_rect=lambda: (None, 1)),
                        ensure_app_ready=Mock(return_value=True))
        with patch.dict(sys.modules, navigation=navigation):
            self.assertFalse(mod.click_pr_menu())
        navigation.activate.assert_called_once_with(1)
        navigation.find_item.assert_not_called()
        navigation.mouse.assert_not_called()

    def test_editor_waits_past_list_description_for_owned_detail(self):
        ticks = itertools.count()
        ctrl = SimpleNamespace(Clock=SimpleNamespace(monotonic=lambda: next(ticks) / 2),
                               checkpoint=Mock(), sleep=Mock(), bind_window=Mock())
        api = Mock()
        api.GetForegroundWindow.side_effect = [1, 2, 2]
        api.GetWindow.return_value = 1
        api.IsWindowEnabled.return_value = True
        mod = functions('pr_form.py', 'editor_window', control=ctrl, win32gui=api,
                        navigation=Mock(), roi_helpers=SimpleNamespace(application_windows=lambda:
                            [(1, 'MNR035 Purchase Requisition List', ''),
                             (2, 'MNR035 Purchase Requisition Detail', '')]))
        self.assertEqual(mod.editor_window(1), 2)
        ctrl.bind_window.assert_called_once_with(2)
        self.assertEqual(ctrl.sleep.call_count, 1)

    def test_list_description_never_counts_as_ready_editor(self):
        ticks = itertools.count()
        ctrl = SimpleNamespace(Clock=SimpleNamespace(monotonic=lambda: next(ticks)),
                               checkpoint=Mock(), sleep=Mock(), bind_window=Mock())
        api = Mock()
        api.GetForegroundWindow.return_value = 1
        mod = functions('pr_form.py', 'editor_window', control=ctrl, win32gui=api,
                        navigation=Mock(), roi_helpers=SimpleNamespace(application_windows=lambda:
                            [(1, 'MNR035 Purchase Requisition List', '')]))
        with self.assertRaises(RuntimeError):
            mod.editor_window(1)
        ctrl.bind_window.assert_not_called()

    def test_monitor_clamp_handles_left_monitor_and_top_edge(self):
        mod = functions('window_position.py', 'clamp')
        self.assertEqual(mod.clamp(-2000, -300, 400, 200, (-1920, 0, 0, 1080)), (-1920, 0))
        self.assertEqual(mod.clamp(-10, 1000, 400, 200, (-1920, 0, 0, 1080)), (-400, 880))

    def test_pr_lookup_prefers_foreground_floating_window(self):
        api = Mock()
        api.EnumChildWindows.side_effect = lambda h, cb, out: None
        api.IsWindowVisible.return_value = True
        api.GetForegroundWindow.return_value = 2
        api.GetWindowRect.side_effect = lambda h: (h, 0, 100, 100)
        mod = functions('roi_helpers.py', 'get_pr_window_rect', win32gui=api,
                        application_windows=lambda: [(1, 'Purchase Request', ''),
                                                     (2, 'MNR035 Purchase Requisition List', '')])
        self.assertEqual(mod.get_pr_window_rect()[1], 2)

    def test_unknown_control_blocks_paste_before_input(self):
        api = Mock()
        mod = functions('pr_form.py', 'text_field', control=Mock(), click=Mock(), point=Mock(),
                        focused_value=lambda h: None, pyautogui=api)
        with self.assertRaises(RuntimeError):
            mod.text_field(1, 'x', 'Description', 'text', 10)
        api.hotkey.assert_not_called()

    def test_mismatched_readback_cannot_succeed(self):
        import re
        ticks = itertools.count()
        timer = SimpleNamespace(monotonic=lambda: next(ticks) / 2)
        mod = functions('pr_form.py', 'verify', re=re,
                        control=SimpleNamespace(Clock=timer, guard=Mock(), sleep=Mock()),
                        focused_value=lambda h: 'different')
        with self.assertRaises(RuntimeError):
            mod.verify(1, 'expected', 'Description')

    def test_editor_cannot_follow_unrelated_foreground(self):
        ticks = itertools.count()
        ctrl = SimpleNamespace(Clock=SimpleNamespace(monotonic=lambda: next(ticks)),
                               checkpoint=Mock(), sleep=Mock(), bind_window=Mock())
        api = Mock()
        api.GetForegroundWindow.return_value = 99
        api.GetWindow.return_value = 0
        mod = functions('pr_form.py', 'editor_window', control=ctrl, win32gui=api,
                        navigation=Mock(), roi_helpers=SimpleNamespace(get_pr_window_rect=lambda: (None, 1),
                                                    application_windows=lambda: [(1, 'PR', '')]))
        with self.assertRaises(RuntimeError):
            mod.editor_window(1)
        ctrl.bind_window.assert_not_called()

    def test_actual_rcc_title(self):
        mod = functions('roi_helpers.py', '_role')
        self.assertEqual(mod._role('Remote Control Center - [ARMGC Monitor]'), 'rcc')

    def test_cancel_blocks_next_input(self):
        event = threading.Event()
        api = Mock()
        with control.session(event, Mock()):
            event.set()
            with self.assertRaises(control.TaskStopped):
                control.Input(api).click(1, 2)
        api.click.assert_not_called()

    def test_focus_change_blocks_input(self):
        api = Mock()
        win = Mock()
        win.GetForegroundWindow.return_value = 2
        with patch.dict(sys.modules, win32gui=win), control.session(threading.Event(), Mock()):
            control.bind_window(1)
            with self.assertRaises(control.TaskStopped):
                control.Input(api).hotkey('ctrl', 'v')
        api.hotkey.assert_not_called()

    def test_wait_is_interruptible(self):
        event = threading.Event()
        with control.session(event, Mock()):
            event.set()
            with self.assertRaises(control.TaskStopped):
                control.sleep(10)

    def test_invalid_form_rejected(self):
        mod = functions('pr_form.py', 'validate', find_index_by_prefix=lambda a: 0 if a == 'known' else None)
        for values in [(' ', 'known', ''), ('desc', 'bad', ''), ('desc', 'known', '123')]:
            with self.assertRaises(ValueError):
                mod.validate(*values)
        mod.validate('desc', 'known', '')

    def test_field_verification_failure_stops_following_fields(self):
        field = Mock(side_effect=RuntimeError('cannot verify Description'))
        mod = functions('pr_form.py', 'fill', control=Mock(), click_add=Mock(), point=Mock(), click=Mock(), text_field=field, editor_window=lambda h: h)
        with self.assertRaisesRegex(RuntimeError, 'Description'):
            mod.fill(1, 'desc', False, 'known', '')
        self.assertEqual(field.call_count, 1)
