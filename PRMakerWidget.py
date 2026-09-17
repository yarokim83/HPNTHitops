"""PR Maker widget: visible task state, cancellation and verified PR entry."""
import os
import queue
import sys
import threading
import time
from pathlib import Path
import dpi_support
dpi_support.enable_per_monitor()
from tkinter import messagebox

import customtkinter as ctk
from PIL import Image
import pyautogui
import pystray
import win32api
import win32con
import win32event
import win32gui

import app_settings as settings
from account_codes import ACCOUNT_CODES
import main
import menu_navigator
import navigation
import pr_form
import task_control as control
import window_position

try:
    from pynput import keyboard
except ImportError:
    keyboard = None

BG = '#17191F'
PANEL = '#242832'
BLUE = '#3985FF'
TEXT = '#EDF1F7'
MUTEX_NAME = r'Local\PRMakerWidget.Instance'
SHOW_EVENT = r'Local\PRMakerWidget.Show'


class Tooltip:
    def __init__(self, widget, text, app):
        self.popup = None
        self.job = None
        self.widget, self.text, self.app = widget, text, app
        widget.bind('<Enter>', self.schedule, add='+')
        widget.bind('<Leave>', self.hide, add='+')
        widget.bind('<Button-1>', self.hide, add='+')

    def schedule(self, event=None):
        self.job = self.widget.after(500, self.show)

    def show(self):
        self.job = None
        if self.app.is_running:
            return
        self.popup = ctk.CTkToplevel(self.widget)
        self.popup.overrideredirect(True)
        self.popup.attributes('-topmost', True)
        ctk.CTkLabel(self.popup, text=self.text, padx=10, pady=5).pack()
        window_position.place(self.popup, self.widget.winfo_rootx(), self.widget.winfo_rooty() + 40)

    def hide(self, event=None):
        if self.job:
            self.widget.after_cancel(self.job)
            self.job = None
        if self.popup:
            self.popup.destroy()
            self.popup = None


class PRMakerWidget(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title('PR Maker ' + settings.VERSION)
        self.overrideredirect(True)
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.attributes('-alpha', 0.98)
        self.is_running = False
        self.pr_visible = False
        self.cancel_event = threading.Event()
        self.events = queue.Queue()
        self.preferences = settings.load()
        self.last_task = None
        self.closing = False
        self.tray = None
        self.listener = None
        self.dialog = None
        self.progress_text = '준비됨 · Ctrl+Shift+P로 열기'
        self.started_at = 0
        self.assets_dir = Path(__file__).parent / 'assets'
        self.icons = []
        self.tips = []
        self.shell = ctk.CTkFrame(self, fg_color=BG, corner_radius=16, border_width=1,
                                  border_color='#454B59')
        self.shell.pack(fill='both', expand=True, padx=1, pady=1)
        self.shell.grid_columnconfigure(0, weight=1)
        dock = ctk.CTkFrame(self.shell, fg_color='transparent')
        dock.grid(row=0, column=0, sticky='ew', padx=8, pady=(6, 0))
        self.task_buttons = []
        for label, image, command, tip in [
            ('PR', 'tools_icon.png', self.toggle_pr_section, '구매요청 입력 펼치기'),
            ('M&C', 'mc_icon.png', lambda: self.run_task('M&C', menu_navigator.run_mc_sequence), 'M&C → Berthing Schedule'),
            ('RCC', 'rcc_icon.png', lambda: self.run_task('RCC', menu_navigator.click_rcc_menu), 'Remote Control Center 열기'),
            ('로그인', None, lambda: self.run_task('Login', menu_navigator.ensure_app_ready), 'HI-TOPS 실행·로그인'),
        ]:
            icon = None
            if image:
                try:
                    icon = ctk.CTkImage(Image.open(self.assets_dir / image), size=(22, 22))
                    self.icons.append(icon)
                except OSError:
                    pass
            btn = ctk.CTkButton(dock, text=label, image=icon, compound='top', width=58, height=50,
                               fg_color=PANEL, hover_color='#364153', command=command)
            btn.pack(side='left', padx=3)
            self.task_buttons.append(btn)
            self.tips.append(Tooltip(btn, tip, self))
        self.settings_btn = ctk.CTkButton(dock, text='⚙', width=38, height=34, command=self.show_settings_dialog)
        self.settings_btn.pack(side='left', padx=(7, 3))
        hide = ctk.CTkButton(dock, text='✕', width=32, height=34, fg_color=PANEL, command=self.withdraw)
        hide.pack(side='left', padx=3)
        self.tips.append(Tooltip(hide, '트레이로 숨기기 · 종료는 트레이 메뉴', self))
        dock.bind('<Button-1>', self.start_move)
        dock.bind('<B1-Motion>', self.do_move)

        self.form = ctk.CTkFrame(self.shell, fg_color='transparent')
        self.form.grid_columnconfigure(0, weight=3)
        self.form.grid_columnconfigure(1, weight=2)
        self.desc_entry = ctk.CTkEntry(self.form, placeholder_text='PR 설명 (필수)', height=34)
        self.desc_entry.grid(row=0, column=0, columnspan=3, sticky='ew', padx=4, pady=4)
        recent = [c for c in self.preferences.get('recent_accounts', []) if c in ACCOUNT_CODES]
        self.account_choices = recent + [c for c in ACCOUNT_CODES if c not in recent]
        self.account_combo = ctk.CTkComboBox(self.form, values=self.account_choices, width=340, height=34)
        self.account_combo.grid(row=1, column=0, sticky='ew', padx=4, pady=4)
        self.account_combo.set(recent[0] if recent else ACCOUNT_CODES[0])
        self.account_combo.bind('<KeyRelease>', self.filter_accounts)
        self.part_entry = ctk.CTkEntry(self.form, placeholder_text='Part No (선택·4자 이상)', height=34)
        self.part_entry.grid(row=1, column=1, sticky='ew', padx=4, pady=4)
        self.unit_price_var = ctk.BooleanVar(value=False)
        self.unit_check = ctk.CTkCheckBox(self.form, text='단가계약', variable=self.unit_price_var, width=90)
        self.unit_check.grid(row=1, column=2, padx=6)
        actions = ctk.CTkFrame(self.form, fg_color='transparent')
        actions.grid(row=2, column=0, columnspan=3, sticky='ew', padx=4, pady=4)
        self.clear_btn = ctk.CTkButton(actions, text='입력 지우기', width=100, fg_color=PANEL, command=self.clear_inputs)
        self.clear_btn.pack(side='left')
        self.run_btn = ctk.CTkButton(actions, text='▶ PR 입력', width=110, fg_color='#207E62', command=self.run_automation_thread)
        self.run_btn.pack(side='right')
        ctk.CTkLabel(actions, text='입력 후 HI-TOPS에서 검토·저장', text_color='#AAB4C7').pack(side='right', padx=12)
        self.status_frame = ctk.CTkFrame(self.shell, fg_color='transparent')
        self.status_frame.grid(row=2, column=0, sticky='ew', padx=12, pady=(2, 8))
        self.status_frame.grid_columnconfigure(0, weight=1)
        self.status_label = ctk.CTkLabel(self.status_frame, text=self.progress_text, anchor='w',
                                        justify='left', wraplength=245, text_color=TEXT)
        self.status_label.grid(row=0, column=0, sticky='ew')
        self.stop_btn = ctk.CTkButton(self.status_frame, text='■ 중지', width=72, state='disabled',
                                     fg_color='#9C3F48', command=self.cancel_task)
        self.stop_btn.grid(row=0, column=1, padx=(5, 0))
        self.retry_btn = ctk.CTkButton(self.status_frame, text='다시 시도', width=82, state='disabled', command=self.retry)
        self.retry_btn.grid(row=1, column=1, pady=(2, 0))
        self.log_btn = ctk.CTkButton(self.status_frame, text='로그 열기', width=85, fg_color=PANEL, command=self.open_log)
        self.log_btn.grid(row=1, column=0, sticky='w', pady=(2, 0))
        self.bind('<Escape>', lambda e: self.cancel_task() if self.is_running else self.withdraw())
        self.bind('<Control-Return>', lambda e: self.run_automation_thread())
        self.protocol('WM_DELETE_WINDOW', self.request_exit)
        self.resize()
        self.after(100, self.poll_events)

    def resize(self):
        self.geometry('760x280' if self.pr_visible else '380x150')
        self.status_label.configure(wraplength=620 if self.pr_visible else 245)
        window_position.place(self, self.winfo_x(), self.winfo_y())

    def toggle_pr_section(self):
        if self.is_running:
            return
        self.pr_visible = not self.pr_visible
        if self.pr_visible:
            self.form.grid(row=1, column=0, sticky='ew', padx=10, pady=3)
        else:
            self.form.grid_remove()
        self.resize()
        if self.pr_visible:
            self.desc_entry.focus_set()

    def filter_accounts(self, event=None):
        query = self.account_combo.get().strip().casefold()
        self.account_combo.configure(values=[c for c in self.account_choices if query in c.casefold()] or self.account_choices)

    def clear_inputs(self):
        self.desc_entry.delete(0, 'end')
        self.part_entry.delete(0, 'end')
        self.desc_entry.focus_set()

    def show_at_cursor(self):
        if self.is_running:
            # Showing/focusing this window would interrupt ERP input ownership.
            return
        self.deiconify()
        x, y = pyautogui.position()
        window_position.place(self, x, y + 20)
        self.lift()
        self.focus_force()
        if self.pr_visible:
            self.desc_entry.focus_set()

    def start_move(self, event):
        self.drag = (event.x_root, event.y_root, self.winfo_x(), self.winfo_y())

    def do_move(self, event):
        if not self.is_running:
            x, y, wx, wy = self.drag
            window_position.place(self, wx + event.x_root - x, wy + event.y_root - y)

    def set_busy(self, busy):
        state = 'disabled' if busy else 'normal'
        for widget in self.task_buttons + [self.settings_btn, self.run_btn, self.clear_btn,
                                          self.desc_entry, self.part_entry, self.account_combo,
                                          self.unit_check, self.log_btn]:
            widget.configure(state=state)
        self.stop_btn.configure(state='normal' if busy else 'disabled')
        self.retry_btn.configure(state='disabled')

    def run_task(self, name, fn):
        if self.is_running:
            return
        if self.dialog and self.dialog.winfo_exists():
            self.progress_text = '설정 창을 닫은 뒤 실행해 주세요.'
            self.dialog.lift()
            return
        try:
            navigation.setup_logging()
        except OSError as exc:
            self.status_label.configure(text='로그 폴더를 열지 못했습니다: ' + str(exc))
            return
        self.cancel_event.clear()
        self.is_running = True
        self.last_task = (name, fn) if name != 'PR' else None
        self.started_at = time.monotonic()
        self.progress_text = name + ' 시작 중 · 중지 Ctrl+Shift+F12'
        self.status_label.configure(text_color=TEXT)
        self.set_busy(True)
        for tip in self.tips:
            tip.hide()
        def worker():
            outcome, detail = 'failed', name + ': 화면 진입을 확인하지 못했습니다. HI-TOPS와 로그를 확인해 주세요.'
            try:
                with control.session(self.cancel_event, lambda text: self.events.put(('progress', text))):
                    navigation.log.info('%s started version=%s', name, settings.VERSION)
                    result = fn()
                    control.checkpoint()
                    if result is True:
                        outcome = 'completed'
                        detail = 'PR 입력값 확인 완료 · HI-TOPS에서 검토·저장해 주세요.' if name == 'PR' else name + ' 완료'
            except control.TaskStopped as exc:
                outcome, detail = 'stopped', str(exc)
            except Exception as exc:
                detail = str(exc)
                navigation.log.exception('%s failed', name)
            finally:
                navigation.log.info('%s %s elapsed=%.2fs', name, outcome, time.monotonic() - self.started_at)
                self.events.put(('done', (outcome, detail)))
        threading.Thread(target=worker, daemon=True).start()

    def poll_events(self):
        try:
            while True:
                kind, data = self.events.get_nowait()
                if kind == 'progress':
                    self.progress_text = data
                elif kind == 'show':
                    self.show_at_cursor()
                elif kind == 'exit':
                    self.request_exit()
                    if not self.is_running:
                        return
                elif kind == 'tray_failed':
                    self.progress_text = '트레이를 시작하지 못했습니다. 이 창에서 계속 사용할 수 있습니다.'
                    self.show_at_cursor()
                elif kind == 'done':
                    self.is_running = False
                    self.set_busy(False)
                    outcome, self.progress_text = data
                    self.status_label.configure(text_color='#AEE1BE' if outcome == 'completed' else '#FFCE91')
                    if outcome != 'completed' and self.last_task:
                        self.retry_btn.configure(state='normal')
                    if self.closing:
                        self.finish_exit()
                        return
        except queue.Empty:
            pass
        if hasattr(self, 'show_event') and win32event.WaitForSingleObject(self.show_event, 0) == 0:
            self.show_at_cursor()
        text = self.progress_text
        if self.is_running:
            if self.cancel_event.is_set():
                text = '중지 요청됨 · 진행 중인 인식 작업 종료 대기'
            text += f' · {int(time.monotonic() - self.started_at)}초'
        self.status_label.configure(text=text)
        self.after(100, self.poll_events)

    def cancel_task(self):
        if self.is_running:
            self.cancel_event.set()
            self.progress_text = '중지 요청됨 · 진행 중인 인식 작업 종료 대기'

    def retry(self):
        if self.last_task and not self.is_running:
            self.run_task(*self.last_task)

    def run_automation_thread(self):
        if self.is_running:
            return
        description = self.desc_entry.get().strip()
        account = self.account_combo.get().strip()
        part = self.part_entry.get().strip()
        try:
            pr_form.validate(description, account, part)
        except ValueError as exc:
            self.progress_text = str(exc)
            self.status_label.configure(text_color='#FFCE91')
            return
        index = next(i for i, c in enumerate(ACCOUNT_CODES) if c.split('/')[0] == account.split('/')[0])
        account = ACCOUNT_CODES[index]
        recent = [account] + [c for c in self.preferences.get('recent_accounts', []) if c != account and c in ACCOUNT_CODES]
        self.preferences['recent_accounts'] = recent[:5]
        self.account_choices = recent[:5] + [c for c in ACCOUNT_CODES if c not in recent[:5]]
        self.account_combo.configure(values=self.account_choices)
        self.account_combo.set(account)
        try:
            settings.save(self.preferences)
        except OSError:
            pass  # Preferences must not prevent entering a PR.
        unit = self.unit_price_var.get()
        self.run_task('PR', lambda: main.run_automation(description, unit, account, part))

    def open_log(self):
        folder = settings.DATA_DIR / 'logs'
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / 'navigation.log'
        os.startfile(str(path if path.exists() else folder))

    def get_startup_path(self):
        return str(Path(os.environ['APPDATA']) / r'Microsoft\Windows\Start Menu\Programs\Startup\PRMakerWidget.lnk')

    def toggle_startup(self, enabled):
        path = self.get_startup_path()
        if enabled:
            import win32com.client
            shortcut = win32com.client.Dispatch('WScript.Shell').CreateShortCut(path)
            executable = str(settings.INSTALL_EXE) if settings.INSTALL_EXE.exists() else sys.executable
            shortcut.TargetPath = executable
            shortcut.Arguments = '' if settings.INSTALL_EXE.exists() or getattr(sys, 'frozen', False) else f'"{Path(__file__).resolve()}"'
            shortcut.WorkingDirectory = str(Path(executable).parent)
            shortcut.save()
        elif os.path.exists(path):
            os.remove(path)

    def show_settings_dialog(self):
        if self.is_running:
            return
        if self.dialog and self.dialog.winfo_exists():
            self.dialog.lift()
            return
        dialog = self.dialog = ctk.CTkToplevel(self)
        dialog.title('PR Maker 설정 · ' + settings.VERSION)
        dialog.geometry('380x360')
        dialog.resizable(False, False)
        dialog.attributes('-topmost', True)
        startup = ctk.BooleanVar(value=os.path.exists(self.get_startup_path()))
        def toggle():
            try:
                self.toggle_startup(startup.get())
            except Exception as exc:
                startup.set(not startup.get())
                messagebox.showerror('시작 프로그램 변경 실패', str(exc), parent=dialog)
        ctk.CTkCheckBox(dialog, text='Windows 로그인 시 자동 실행', variable=startup, command=toggle).pack(padx=18, pady=18, anchor='w')
        ctk.CTkLabel(dialog, text='HI-TOPS 암호 변경', font=('맑은 고딕', 15, 'bold')).pack()
        current = ctk.CTkFrame(dialog, fg_color='transparent')
        current.pack(padx=35, pady=(8, 4), fill='x')
        ctk.CTkLabel(current, text='현재 암호', width=68).pack(side='left')
        current_entry = ctk.CTkEntry(current, width=242)
        current_entry.pack(side='left', fill='x', expand=True)
        current_entry.insert(0, menu_navigator.get_password())
        current_entry.configure(state='disabled')
        entry = ctk.CTkEntry(dialog, placeholder_text='새 암호', show='•', width=310)
        entry.pack(pady=8)
        confirm = ctk.CTkEntry(dialog, placeholder_text='새 암호 확인', show='•', width=310)
        confirm.pack(pady=4)
        visible = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(dialog, text='입력한 암호 보기', variable=visible,
                       command=lambda: [e.configure(show='' if visible.get() else '•') for e in (entry, confirm)]).pack(padx=35, pady=8, anchor='w')
        def save():
            if not entry.get() or entry.get() != confirm.get():
                messagebox.showwarning('암호 확인', '두 칸에 같은 새 암호를 입력해 주세요.', parent=dialog)
                return
            try:
                menu_navigator.save_password(entry.get())
            except (OSError, ValueError, TypeError) as exc:
                messagebox.showerror('암호 저장 실패', str(exc), parent=dialog)
                return
            self.progress_text = '암호를 저장했습니다.'
            dialog.destroy()
        ctk.CTkButton(dialog, text='암호 저장', command=save).pack(pady=8)
        dialog.bind('<Escape>', lambda e: dialog.destroy())
        dialog.after(50, lambda: window_position.place(dialog, self.winfo_x(), self.winfo_y() - 370))

    def request_exit(self):
        self.closing = True
        if self.is_running:
            self.cancel_task()
        else:
            self.finish_exit()

    def finish_exit(self):
        if self.tray:
            self.tray.stop()
        if self.listener:
            self.listener.stop()
        self.destroy()


def setup_tray(app):
    try:
        app.tray = pystray.Icon('PRMaker', Image.open(app.assets_dir / 'taskbar_icon.png'),
                               'PR Maker ' + settings.VERSION, pystray.Menu(
            pystray.MenuItem('PR Maker 열기', lambda *a: app.events.put(('show', None)), default=True),
            pystray.MenuItem('자동화 중지', lambda *a: app.cancel_event.set()),
            pystray.MenuItem('종료', lambda *a: app.events.put(('exit', None)))))
        app.tray.run()
    except Exception:
        app.events.put(('tray_failed', None))


def start():
    mutex = win32event.CreateMutex(None, False, MUTEX_NAME)
    duplicate = win32api.GetLastError() == 183
    event = win32event.CreateEvent(None, False, False, SHOW_EVENT)
    if duplicate:
        win32event.SetEvent(event)
        win32api.CloseHandle(event)
        win32api.CloseHandle(mutex)
        return
    try:
        ctk.set_appearance_mode('dark')
        app = PRMakerWidget()
        app.show_event = event
        if keyboard:
            app.listener = keyboard.GlobalHotKeys({
                '<ctrl>+<shift>+p': lambda: app.events.put(('show', None)),
                '<ctrl>+<shift>+<f12>': app.cancel_event.set,
            })
            try:
                app.listener.start()
            except Exception:
                app.listener = None
                app.progress_text = '단축키를 시작하지 못했습니다. 트레이 메뉴로 열어 주세요.'
        app.withdraw()
        threading.Thread(target=setup_tray, args=(app,), daemon=True).start()
        app.mainloop()
    finally:
        win32api.CloseHandle(event)
        win32api.CloseHandle(mutex)


if __name__ == '__main__':
    if '--self-check' in sys.argv:
        # Packaging smoke check: imports only, no windows, hotkeys, or ERP actions.
        sys.exit(0)
    start()
