import itertools
import sys
import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch
from navigation_regression import functions


class LoginTests(unittest.TestCase):
    def test_main_is_found_behind_lingering_login(self):
        api = Mock()
        api.GetClassName.return_value = 'WindowsForms10.Window'
        mod = functions('login_manager.py', 'main_window', win32gui=api,
                        is_login=lambda h: h == 1,
                        roi_helpers=SimpleNamespace(application_windows=lambda: [
                            (1, 'HI-TOPS III Login', ''),
                            (3, 'Maintenance & Repair System', ''),
                            (2, 'HiTOPS3', '')]))
        self.assertEqual(mod.main_window(), 2)

    def test_slow_login_transition_waits_without_input(self):
        clock = itertools.count()
        api = Mock()
        mod = functions('login_manager.py', 'wait_main', control=Mock(),
                        time=SimpleNamespace(monotonic=lambda: next(clock), sleep=Mock()),
                        main_window=Mock(side_effect=[None, None, None, 2]), win32gui=api)
        self.assertTrue(mod.wait_main())
        self.assertEqual(mod.time.sleep.call_count, 3)

    def test_activation_failure_can_recover_via_main_without_typing(self):
        activate = Mock(return_value=False)
        api = Mock()
        inputs = Mock()
        mod = functions('login_manager.py', 'perform_login', control=Mock(),
                        time=SimpleNamespace(monotonic=lambda: 0, sleep=Mock()),
                        main_window=lambda: None, wait_main=Mock(return_value=True),
                        roi_helpers=SimpleNamespace(application_windows=lambda: [(1,'Login','')]),
                        is_login=lambda h: True, win32gui=api, pyautogui=inputs)
        with patch.dict(sys.modules, menu_navigator=SimpleNamespace(force_activate_window=activate)):
            self.assertTrue(mod.perform_login('TEST-ONLY'))
        self.assertEqual(activate.call_count, 3)
        inputs.write.assert_not_called()
        mod.wait_main.assert_called_once()

    def test_submit_once_then_wait_on_delayed_main(self):
        activate = Mock(return_value=True)
        api = Mock()
        api.GetForegroundWindow.return_value = 1
        api.GetWindowRect.return_value = (0,0,400,300)
        inputs = Mock()
        mod = functions('login_manager.py', 'perform_login', control=Mock(),
                        time=SimpleNamespace(monotonic=lambda: 0, sleep=Mock()),
                        main_window=lambda: None, wait_main=Mock(return_value=False),
                        roi_helpers=SimpleNamespace(application_windows=lambda: [(1,'Login','')]),
                        is_login=lambda h: True, win32gui=api, pyautogui=inputs)
        with patch.dict(sys.modules, menu_navigator=SimpleNamespace(force_activate_window=activate)):
            self.assertFalse(mod.perform_login('TEST-ONLY'))
        inputs.write.assert_called_once()
        inputs.press.assert_called_once_with('enter')
        mod.wait_main.assert_called_once()

    def test_wait_timeout_does_not_claim_success(self):
        clock=itertools.count()
        mod=functions('login_manager.py','wait_main', control=Mock(),
                      time=SimpleNamespace(monotonic=lambda: next(clock), sleep=Mock()),
                      main_window=lambda: None, win32gui=Mock(), roi_helpers=Mock())
        self.assertFalse(mod.wait_main(timeout=3))
