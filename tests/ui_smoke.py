"""Create only hidden test widgets; never open or drive the ERP."""
import sys
from pathlib import Path
from unittest.mock import Mock, patch
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import PRMakerWidget as ui
import customtkinter as ctk

original = ctk.CTk.__init__
def hidden(self, *args, **kwargs):
    original(self, *args, **kwargs)
    self.withdraw()

with patch.object(ctk.CTk, '__init__', hidden), patch.object(ui.settings, 'load', return_value={}):
    app = ui.PRMakerWidget()
    try:
        app.update_idletasks()
        scale = app.shell._get_widget_scaling()
        print('DPI scale:', scale)
        app.toggle_pr_section()
        app.update_idletasks()
        app.resize()
        assert app.account_combo.grid_info()['columnspan'] == 2
        assert app.status_label.grid_info()['columnspan'] == 3
        # Simulate independent widget/window scale changes, including the mismatch
        # that previously clipped the bottom of the fixed-height window.
        for widget_factor, window_factor in ((1, 1), (1.25, 1), (1.5, 1.5), (2, 1.5), (2.4, 2.4), (1, 1)):
            ctk.set_widget_scaling(widget_factor)
            ctk.set_window_scaling(window_factor)
            app.update_idletasks()
            app.resize()
            app.update_idletasks()
            physical_height = app._current_height * app._get_window_scaling()
            assert physical_height > app.status_frame.winfo_reqheight() + 60
            assert app.status_frame.winfo_reqwidth() <= app._current_width * app._get_window_scaling()
        print('Scale-transition layout checks passed')
        app.desc_entry.insert(0, ' ')
        with patch.object(app, 'run_task') as run:
            app.run_automation_thread()
            run.assert_not_called()
        app.is_running = True
        with patch.object(app, 'deiconify') as show:
            app.show_at_cursor()
            show.assert_not_called()
        app.set_busy(True)
        assert app.run_btn.cget('state') == 'disabled'
        assert app.stop_btn.cget('state') == 'normal'
        app.cancel_task()
        assert app.cancel_event.is_set()
        app.events.put(('done', ('stopped', 'test stopped')))
        app.poll_events()
        assert not app.is_running
        assert app.run_btn.cget('state') == 'normal'
        assert app.retry_btn.cget('state') == 'disabled'
        with patch.object(ui.menu_navigator, 'get_password', return_value='CURRENT-DEMO'):
            with patch.object(ctk.CTkToplevel, 'deiconify'):
                app.show_settings_dialog()
                app.dialog.withdraw()
                app.dialog.update_idletasks()
                entries = []
                def collect(widget):
                    if isinstance(widget, ctk.CTkEntry):
                        entries.append(widget)
                    for child in widget.winfo_children():
                        collect(child)
                collect(app.dialog)
                current = next(entry for entry in entries if entry.get() == 'CURRENT-DEMO')
                assert current.cget('show') == ''
                assert current.cget('state') == 'disabled'
                assert app.dialog.winfo_reqheight() <= 400 * app.dialog._get_window_scaling()
                app.dialog.destroy()
        app.events.put(('exit', None))
        with patch.object(app, 'finish_exit') as exit_app, patch.object(app, 'after') as reschedule:
            app.poll_events()
            exit_app.assert_called_once()
            reschedule.assert_not_called()
        print('Hidden UI state checks passed')
    finally:
        app.destroy()
