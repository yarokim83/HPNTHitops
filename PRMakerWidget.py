import customtkinter as ctk
import pyautogui
import threading
import sys
import os
import time
import main
import menu_navigator
import ocr_helpers

# Check for pynput
try:
    from pynput import keyboard
    PYNPUT_AVAILABLE = True
except ImportError:
    PYNPUT_AVAILABLE = False
    print("Warning: 'pynput' library not found. Global hotkey will not work.")

# ── Apple-style Widget Configuration ──
WIDGET_WIDTH_FULL = 820
WIDGET_WIDTH_MINI = 340
WIDGET_HEIGHT = 72
ALPHA_VALUE = 0.92

# ── Apple Color Palette ──
COLORS = {
    "bg":            "#1C1C1E",      # iOS dark background
    "bg_secondary":  "#2C2C2E",      # Card/surface
    "bg_tertiary":   "#3A3A3C",      # Hover state
    "accent_blue":   "#0A84FF",      # iOS blue
    "accent_green":  "#30D158",      # iOS green
    "accent_red":    "#FF453A",      # iOS red
    "accent_orange": "#FF9F0A",      # iOS orange
    "text_primary":  "#FFFFFF",      # Primary text
    "text_secondary":"#8E8E93",      # Secondary text
    "separator":     "#48484A",      # Separator/divider
    "hover":         "#3A3A3C",      # Button hover
    "pressed":       "#545456",      # Button pressed
}


class PRMakerWidget(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Assets
        self.assets_dir = os.path.join(os.path.dirname(__file__), 'assets')

        # Initialize OCR Engine
        threading.Thread(target=ocr_helpers.init_tesseract, daemon=True).start()

        # ── Window Setup ──
        self.title("PR Maker")
        self.geometry(f"{WIDGET_WIDTH_MINI}x{WIDGET_HEIGHT}")
        self.overrideredirect(True)  # Frameless
        self.attributes('-alpha', ALPHA_VALUE)
        self.resizable(False, False)
        self.configure(fg_color=COLORS["bg"])

        # Center initially
        self.center_window()

        # ── Main Container with rounded appearance ──
        self.main_container = ctk.CTkFrame(
            self, corner_radius=18, fg_color=COLORS["bg"],
            border_width=1, border_color=COLORS["separator"]
        )
        self.main_container.pack(fill="both", expand=True, padx=1, pady=1)
        self.main_container.grid_columnconfigure(0, weight=0)  # Sidebar
        self.main_container.grid_columnconfigure(1, weight=0)  # Divider
        self.main_container.grid_columnconfigure(2, weight=1)  # Inputs
        self.main_container.grid_rowconfigure(0, weight=1)

        # ── Dock Bar (Icon Buttons) ──
        self.dock = ctk.CTkFrame(self.main_container, corner_radius=14, fg_color=COLORS["bg_secondary"])
        self.dock.grid(row=0, column=0, padx=6, pady=6, sticky="ns")

        # Load Images
        try:
            from PIL import Image
            icon_size = (36, 36)
            self.img_tools = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "tools_icon.png")), size=icon_size)
            self.img_mc    = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "mc_icon.png")),    size=icon_size)
            self.img_rcc   = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "rcc_icon.png")),   size=icon_size)
        except Exception as e:
            print(f"Failed to load icons: {e}")
            self.img_tools = self.img_mc = self.img_rcc = None

        # ── Dock Buttons (Apple-style pill buttons) ──
        btn_size = 52
        btn_cfg = dict(
            width=btn_size, height=btn_size,
            corner_radius=14,
            fg_color="transparent",
            hover_color=COLORS["hover"],
            text="",
            border_width=0,
        )

        self.btn_tools = ctk.CTkButton(self.dock, image=self.img_tools,
                                       command=self.toggle_pr_section, **btn_cfg)
        self.btn_tools.pack(side="left", padx=4, pady=6)

        self.btn_mc = ctk.CTkButton(self.dock, image=self.img_mc,
                                    command=lambda: threading.Thread(
                                        target=menu_navigator.run_mc_sequence, daemon=True).start(),
                                    **btn_cfg)
        self.btn_mc.pack(side="left", padx=2, pady=6)

        self.btn_rcc = ctk.CTkButton(self.dock, image=self.img_rcc,
                                     command=lambda: threading.Thread(
                                         target=menu_navigator.click_rcc_menu, daemon=True).start(),
                                     **btn_cfg)
        self.btn_rcc.pack(side="left", padx=2, pady=6)

        # ── Separator ──
        self.sep_line = ctk.CTkFrame(self.dock, width=1, height=32,
                                     corner_radius=0, fg_color=COLORS["separator"])
        self.sep_line.pack(side="left", padx=4, pady=14)

        # ── Settings Button (gear icon) ──
        self.btn_settings = ctk.CTkButton(
            self.dock, text="⚙", width=32, height=32,
            corner_radius=16, font=("SF Pro Display", 16),
            fg_color="transparent", hover_color=COLORS["hover"],
            text_color=COLORS["text_secondary"],
            command=self.show_settings_dialog
        )
        self.btn_settings.pack(side="left", padx=2, pady=6)

        # ── Close Button (subtle X) ──
        self.btn_exit = ctk.CTkButton(
            self.dock, text="✕", width=32, height=32,
            corner_radius=16, font=("SF Pro Display", 14),
            fg_color="transparent", hover_color=COLORS["accent_red"],
            text_color=COLORS["text_secondary"],
            command=self.destroy
        )
        self.btn_exit.pack(side="left", padx=4, pady=6)

        # ── Drag Handle (invisible but functional) ──
        # Bind drag to the entire dock for easier dragging
        self.dock.bind("<Button-1>", self.start_move)
        self.dock.bind("<B1-Motion>", self.do_move)
        self.main_container.bind("<Button-1>", self.start_move)
        self.main_container.bind("<B1-Motion>", self.do_move)

        # ── PR Input Section (Hidden initially) ──
        self.pr_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.pr_frame.grid_columnconfigure(0, weight=2)
        self.pr_frame.grid_columnconfigure(1, weight=0)
        self.pr_frame.grid_columnconfigure(2, weight=3)
        self.pr_frame.grid_columnconfigure(3, weight=1)
        self.pr_frame.grid_columnconfigure(4, weight=0)
        self.pr_frame.grid_columnconfigure(5, weight=0)
        self.pr_frame.grid_rowconfigure(0, weight=1)

        # ── Input Fields (Apple-style) ──
        entry_cfg = dict(
            corner_radius=10,
            border_width=1,
            border_color=COLORS["separator"],
            fg_color=COLORS["bg_secondary"],
            text_color=COLORS["text_primary"],
            font=("SF Pro Display", 13),
            height=34,
        )

        self.desc_entry = ctk.CTkEntry(
            self.pr_frame, placeholder_text="PR Description",
            placeholder_text_color=COLORS["text_secondary"],
            width=160, **entry_cfg
        )
        self.desc_entry.grid(row=0, column=0, padx=(6, 3), pady=10, sticky="ew")

        # Checkbox (modern toggle style)
        self.unit_price_var = ctk.BooleanVar(value=False)
        self.unit_price_chk = ctk.CTkCheckBox(
            self.pr_frame, text="단가",
            variable=self.unit_price_var, width=50,
            font=("SF Pro Display", 12),
            text_color=COLORS["text_secondary"],
            fg_color=COLORS["accent_blue"],
            hover_color=COLORS["bg_tertiary"],
            border_color=COLORS["separator"],
            corner_radius=6,
            checkbox_height=18, checkbox_width=18,
        )
        self.unit_price_chk.grid(row=0, column=1, padx=3, pady=10)

        # Account Code Combo
        self.account_codes = [
            "0501030000/수선유지비", "0501030100/수선유지비", "0501030101/장비 자재비-QC",
            "0501030102/장비 자재비-ATC", "0501030103/장비 자재비-RS", "0501030104/장비 자재비-YT",
            "0501030105/장비 자재비-YC", "0501030106/장비 자재비-FL", "0501030107/장비 자재비-기타",
            "0501030108/수선유지비-외주수리-QC", "0501030109/수선유지비-외주수리-ATC",
            "0501030110/수선유지비-외주수리-RS", "0501030111/수선유지비-외주수리-YT",
            "0501030112/수선유지비-외주수리-YC", "0501030113/수선유지비-외주수리-FL",
            "0501030114/수선유지비-외주수리-기타", "0501030115/시설물-야드시설물(자재)",
            "0501030116/수선유지비-시설물-CFS시설물", "0501030117/시설물-전기시설물(자재)",
            "0501030118/시설물-외주수리", "0501030119/수선유지비_작업공구-야드공구",
            "0501030120/수선유지비_작업공구-정비공구", "0501030121/수선유지비_작업공구-CFS공구",
            "0501030122/수선유지비_작업공구-안전공구", "0501030123/수선유지비_작업공구-기타공구",
            "0501030124/수선유지비_작업소모품-야드소모품", "0501030125/작업소모품-정비소모품/공구",
            "0501030126/수선유지비_작업소모품-CFS소모품", "0501030127/수선유지비_작업소모품-안전소모품",
            "0501030128/수선유지비_작업소모품-기타소모품", "0501030129/수선유지비-CNTR",
            "0501030130/수선유지비-기타 (사고변상금등)", "0501030131/장비자재비-ECH",
            "0501030132/수선유지비-외주수리-ECH", "0501040106/동력비-윤활유"
        ]
        self.account_combo = ctk.CTkComboBox(
            self.pr_frame, values=self.account_codes, width=160,
            corner_radius=10,
            border_width=1,
            border_color=COLORS["separator"],
            fg_color=COLORS["bg_secondary"],
            button_color=COLORS["bg_tertiary"],
            button_hover_color=COLORS["hover"],
            text_color=COLORS["text_primary"],
            font=("SF Pro Display", 12),
            dropdown_fg_color=COLORS["bg_secondary"],
            dropdown_text_color=COLORS["text_primary"],
            dropdown_hover_color=COLORS["accent_blue"],
            height=34,
        )
        self.account_combo.grid(row=0, column=2, padx=3, pady=10, sticky="ew")
        self.account_combo.set("0501030100/수선유지비")

        # Part No Input
        self.part_entry = ctk.CTkEntry(
            self.pr_frame, placeholder_text="Part No",
            placeholder_text_color=COLORS["text_secondary"],
            width=80, **entry_cfg
        )
        self.part_entry.grid(row=0, column=3, padx=3, pady=10, sticky="ew")

        # Run Button (Apple-green pill)
        self.run_btn = ctk.CTkButton(
            self.pr_frame, text="▶", width=36, height=34,
            corner_radius=10,
            fg_color=COLORS["accent_green"],
            hover_color="#28B84C",
            text_color="white",
            font=("SF Pro Display", 16),
            command=self.run_automation_thread
        )
        self.run_btn.grid(row=0, column=4, padx=3, pady=10)

        # Close Section Button (subtle)
        self.close_btn = ctk.CTkButton(
            self.pr_frame, text="‹", width=24, height=34,
            corner_radius=10,
            fg_color="transparent",
            hover_color=COLORS["hover"],
            text_color=COLORS["text_secondary"],
            font=("SF Pro Display", 18),
            command=self.toggle_pr_section
        )
        self.close_btn.grid(row=0, column=5, padx=(2, 6), pady=10)

        # Bind Escape to Hide
        self.bind("<Escape>", lambda e: self.hide_widget())

        # Internal State
        self.is_running = False
        self.pr_visible = False

    def toggle_pr_section(self):
        if self.pr_visible:
            self.pr_frame.grid_forget()
            self.geometry(f"{WIDGET_WIDTH_MINI}x{WIDGET_HEIGHT}")
            self.btn_tools.configure(fg_color="transparent")
            self.pr_visible = False
        else:
            self.pr_frame.grid(row=0, column=2, sticky="nsew", padx=(0, 4))
            self.geometry(f"{WIDGET_WIDTH_FULL}x{WIDGET_HEIGHT}")
            self.btn_tools.configure(fg_color=COLORS["accent_blue"])
            self.pr_visible = True
            self.desc_entry.focus_set()


    def show_settings_dialog(self):
        """Apple-style floating settings dialog."""
        dialog = ctk.CTkToplevel(self)
        dialog.title("")
        dialog.overrideredirect(True)
        dialog.attributes('-topmost', True)
        dialog.configure(fg_color=COLORS["bg"])

        # Position above the widget
        wx = self.winfo_x()
        wy = self.winfo_y()
        dw, dh = 280, 240
        dialog.geometry(f"{dw}x{dh}+{wx}+{wy - dh - 8}")

        # Container with border
        container = ctk.CTkFrame(dialog, corner_radius=14,
                                 fg_color=COLORS["bg_secondary"],
                                 border_width=1, border_color=COLORS["separator"])
        container.pack(fill="both", expand=True, padx=2, pady=2)

        # Title
        title_label = ctk.CTkLabel(container, text="설정",
                                   font=("SF Pro Display", 14, "bold"),
                                   text_color=COLORS["text_primary"])
        title_label.pack(pady=(12, 6))

        # ── Startup Toggle ──
        startup_frame = ctk.CTkFrame(container, fg_color="transparent")
        startup_frame.pack(fill="x", padx=16, pady=(0, 10))
        
        is_startup = self.is_in_startup()
        startup_var = ctk.BooleanVar(value=is_startup)
        
        def on_startup_toggle():
            self.toggle_startup(startup_var.get())

        ctk.CTkCheckBox(startup_frame, text="Windows 시작 시 자동 실행",
                        variable=startup_var, command=on_startup_toggle,
                        font=("SF Pro Display", 12),
                        text_color=COLORS["text_primary"],
                        fg_color=COLORS["accent_green"],
                        hover_color=COLORS["bg_tertiary"],
                        border_color=COLORS["separator"],
                        corner_radius=6,
                        checkbox_height=18, checkbox_width=18).pack(side="left")

        # ── Separator ──
        ctk.CTkFrame(container, height=1, fg_color=COLORS["separator"]).pack(fill="x", padx=16, pady=4)

        # ── Password Section ──
        ctk.CTkLabel(container, text="비밀번호 변경",
                     font=("SF Pro Display", 12, "bold"),
                     text_color=COLORS["text_secondary"]).pack(pady=(8, 4))

        # Current password (read-only, masked)
        current_pw = menu_navigator.get_password()
        cur_frame = ctk.CTkFrame(container, fg_color="transparent")
        cur_frame.pack(fill="x", padx=16, pady=2)
        ctk.CTkLabel(cur_frame, text="현재", width=36,
                     font=("SF Pro Display", 11),
                     text_color=COLORS["text_secondary"]).pack(side="left")
        cur_entry = ctk.CTkEntry(cur_frame, height=30, corner_radius=8,
                                 fg_color=COLORS["bg"], border_width=1,
                                 border_color=COLORS["separator"],
                                 text_color=COLORS["text_secondary"],
                                 font=("SF Pro Display", 12))
        cur_entry.pack(side="left", fill="x", expand=True, padx=(4, 0))
        cur_entry.insert(0, current_pw)
        cur_entry.configure(state="disabled")

        # New password input
        new_frame = ctk.CTkFrame(container, fg_color="transparent")
        new_frame.pack(fill="x", padx=16, pady=2)
        ctk.CTkLabel(new_frame, text="변경", width=36,
                     font=("SF Pro Display", 11),
                     text_color=COLORS["text_secondary"]).pack(side="left")
        new_entry = ctk.CTkEntry(new_frame, height=30, corner_radius=8,
                                 fg_color=COLORS["bg"], border_width=1,
                                 border_color=COLORS["separator"],
                                 text_color=COLORS["text_primary"],
                                 placeholder_text="새 비밀번호",
                                 placeholder_text_color=COLORS["text_secondary"],
                                 font=("SF Pro Display", 12))
        new_entry.pack(side="left", fill="x", expand=True, padx=(4, 0))
        # new_entry.focus_set() # Don't auto-focus, let user click if they want to change pw

        # Buttons
        btn_frame = ctk.CTkFrame(container, fg_color="transparent")
        btn_frame.pack(fill="x", padx=16, pady=(12, 12))

        def do_save():
            new_pw = new_entry.get().strip()
            if new_pw:
                menu_navigator.save_password(new_pw)
                print(f"Password updated successfully.")
            dialog.destroy()

        ctk.CTkButton(btn_frame, text="닫기", width=60, height=28,
                      corner_radius=8, fg_color=COLORS["bg_tertiary"],
                      hover_color=COLORS["hover"],
                      text_color=COLORS["text_primary"],
                      font=("SF Pro Display", 12),
                      command=dialog.destroy).pack(side="right", padx=(4, 0))

        ctk.CTkButton(btn_frame, text="저장", width=60, height=28,
                      corner_radius=8, fg_color=COLORS["accent_blue"],
                      hover_color="#0070E0",
                      text_color="white",
                      font=("SF Pro Display", 12, "bold"),
                      command=do_save).pack(side="right")

        # Close on Escape
        dialog.bind("<Escape>", lambda e: dialog.destroy())
        new_entry.bind("<Return>", lambda e: do_save())

    def get_startup_path(self):
        return os.path.join(os.getenv('APPDATA'), r'Microsoft\Windows\Start Menu\Programs\Startup', 'PRMakerWidget.lnk')

    def is_in_startup(self):
        return os.path.exists(self.get_startup_path())

    def toggle_startup(self, enable):
        shortcut_path = self.get_startup_path()
        if enable:
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                shortcut.Targetpath = sys.executable if getattr(sys, 'frozen', False) else sys.executable
                if getattr(sys, 'frozen', False):
                     shortcut.Targetpath = sys.executable
                else:
                     # Running as script: pythonw.exe PRMakerWidget.py
                     # But better to point to a bat file or just the python executable with arguments
                     shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
                
                shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
                shortcut.IconLocation = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets', 'taskbar_icon.png')
                shortcut.save()
                print(f"Startup shortcut created at {shortcut_path}")
            except Exception as e:
                print(f"Failed to create startup shortcut: {e}")
        else:
            if os.path.exists(shortcut_path):
                try:
                    os.remove(shortcut_path)
                    print("Startup shortcut removed.")
                except Exception as e:
                    print(f"Failed to remove startup shortcut: {e}")


    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - WIDGET_WIDTH_MINI) // 2
        y = screen_height - WIDGET_HEIGHT - 80  # Near bottom, like macOS dock
        self.geometry(f"+{x}+{y}")

    def bring_to_front(self):
        self.deiconify()
        self.attributes('-topmost', True)
        self.lift()
        self.focus_force()
        self.after(200, lambda: self.attributes('-topmost', False))

    def start_move(self, event):
        self._drag_x = event.x
        self._drag_y = event.y

    def do_move(self, event):
        deltax = event.x - self._drag_x
        deltay = event.y - self._drag_y
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        self.geometry(f"+{x}+{y}")

    def show_at_cursor(self):
        try:
            mouse_x, mouse_y = pyautogui.position()
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()

            w = WIDGET_WIDTH_FULL if self.pr_visible else WIDGET_WIDTH_MINI
            final_x = mouse_x
            final_y = mouse_y + 20

            if final_x + w > screen_w:
                final_x = screen_w - w - 10
            if final_y + WIDGET_HEIGHT > screen_h:
                final_y = mouse_y - WIDGET_HEIGHT - 10

            self.geometry(f"+{final_x}+{final_y}")
            self.bring_to_front()
            if self.pr_visible:
                self.desc_entry.focus_set()
        except:
            self.center_window()
            self.bring_to_front()

    def hide_widget(self):
        self.withdraw()

    def run_automation_thread(self):
        if self.is_running:
            return
        desc = self.desc_entry.get()
        if not desc:
            # Flash the entry border red briefly
            self.desc_entry.configure(border_color=COLORS["accent_red"])
            self.after(1500, lambda: self.desc_entry.configure(border_color=COLORS["separator"]))
            return
        account = self.account_combo.get()
        part_no = self.part_entry.get()
        self.run_btn.configure(state="disabled", fg_color=COLORS["bg_tertiary"])
        self.is_running = True
        threading.Thread(target=self._run_automation, args=(desc, account, part_no), daemon=True).start()

    def _run_automation(self, desc, account, part_no):
        try:
            is_unit_price = self.unit_price_var.get()
            main.run_automation(desc, is_unit_price, account, part_no)
        except Exception as e:
            print(f"Automation Error: {e}")
        finally:
            self.is_running = False
            self.run_btn.configure(state="normal", fg_color=COLORS["accent_green"])


app = None

def on_activate():
    if app:
        app.after(0, app.show_at_cursor)

def start_hotkey_listener():
    if not PYNPUT_AVAILABLE:
        return
    with keyboard.GlobalHotKeys({'<ctrl>+<shift>+p': on_activate}) as h:
        h.join()

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    app = PRMakerWidget()
    print("=" * 50)
    print(" PR Maker Widget — Apple Style")
    print("=" * 50)
    if PYNPUT_AVAILABLE:
        print(" [Ctrl + Shift + P] to show widget at mouse cursor")
        threading.Thread(target=start_hotkey_listener, daemon=True).start()
    else:
        print(" [!] pynput not installed. Global hotkey disabled.")
    app.mainloop()
