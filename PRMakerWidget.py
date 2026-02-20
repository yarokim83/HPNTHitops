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

# ── Liquid Glass Widget Configuration ──
WIDGET_WIDTH_FULL = 840
WIDGET_WIDTH_MINI = 300
WIDGET_HEIGHT = 68
ALPHA_VALUE = 0.88

# ── Apple Liquid Glass Color System ──
# Inspired by iOS 26 / macOS Tahoe Liquid Glass
GLASS = {
    # Glass surfaces
    "glass_bg":         "#0D0D0F",      # Ultra-dark glass base
    "glass_surface":    "#1A1A1F",      # Frosted glass panel
    "glass_elevated":   "#252530",      # Elevated glass card
    "glass_hover":      "#2E2E3A",      # Hover state
    "glass_pressed":    "#383845",      # Pressed state

    # Liquid accents (vibrant through glass)
    "liquid_blue":      "#007AFF",      # Primary action
    "liquid_cyan":      "#5AC8FA",      # Accent/highlight
    "liquid_green":     "#30D158",      # Success/Run
    "liquid_green_dim": "#1B8A3A",      # Green hover
    "liquid_red":       "#FF453A",      # Danger/Close
    "liquid_orange":    "#FF9F0A",      # Warning
    "liquid_purple":    "#BF5AF2",      # Special

    # Typography
    "text_bright":      "#F5F5F7",      # Primary text (Apple white)
    "text_mid":         "#A1A1A6",      # Secondary text
    "text_dim":         "#636366",      # Tertiary/placeholder

    # Structure
    "border_glass":     "#38383D",      # Glass border (subtle)
    "border_glow":      "#4A4A52",      # Active border glow
    "divider":          "#2C2C30",      # Separator line
}


class PRMakerWidget(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Assets
        self.assets_dir = os.path.join(os.path.dirname(__file__), 'assets')

        # Initialize OCR Engine
        threading.Thread(target=ocr_helpers.init_tesseract, daemon=True).start()

        # ── Frameless Glass Window ──
        self.title("PR Maker")
        self.geometry(f"{WIDGET_WIDTH_MINI}x{WIDGET_HEIGHT}")
        self.overrideredirect(True)
        self.attributes('-alpha', ALPHA_VALUE)
        self.resizable(False, False)
        self.configure(fg_color=GLASS["glass_bg"])

        # Center near bottom (dock-style)
        self.center_window()

        # ── Outer Glass Shell ──
        self.glass_shell = ctk.CTkFrame(
            self, corner_radius=22, fg_color=GLASS["glass_surface"],
            border_width=1, border_color=GLASS["border_glass"]
        )
        self.glass_shell.pack(fill="both", expand=True, padx=1, pady=1)
        self.glass_shell.grid_columnconfigure(0, weight=0)  # Dock
        self.glass_shell.grid_columnconfigure(1, weight=1)  # Content
        self.glass_shell.grid_rowconfigure(0, weight=1)

        # ── Glass Dock Bar ──
        self.dock = ctk.CTkFrame(
            self.glass_shell, corner_radius=16,
            fg_color="transparent"
        )
        self.dock.grid(row=0, column=0, padx=(6, 2), pady=5, sticky="ns")

        # Load Icons
        try:
            from PIL import Image
            icon_size = (32, 32)
            self.img_tools = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "tools_icon.png")), size=icon_size)
            self.img_mc    = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "mc_icon.png")),    size=icon_size)
            self.img_rcc   = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "rcc_icon.png")),   size=icon_size)
        except Exception as e:
            print(f"Failed to load icons: {e}")
            self.img_tools = self.img_mc = self.img_rcc = None

        # ── Dock Buttons (Glass Pill Buttons) ──
        btn_size = 46
        glass_btn = dict(
            width=btn_size, height=btn_size,
            corner_radius=14,
            fg_color="transparent",
            hover_color=GLASS["glass_hover"],
            text="",
            border_width=0,
        )

        self.btn_tools = ctk.CTkButton(
            self.dock, image=self.img_tools,
            command=self.toggle_pr_section, **glass_btn
        )
        self.btn_tools.pack(side="left", padx=3, pady=4)

        self.btn_mc = ctk.CTkButton(
            self.dock, image=self.img_mc,
            command=lambda: threading.Thread(
                target=menu_navigator.run_mc_sequence, daemon=True).start(),
            **glass_btn
        )
        self.btn_mc.pack(side="left", padx=2, pady=4)

        self.btn_rcc = ctk.CTkButton(
            self.dock, image=self.img_rcc,
            command=lambda: threading.Thread(
                target=menu_navigator.click_rcc_menu, daemon=True).start(),
            **glass_btn
        )
        self.btn_rcc.pack(side="left", padx=2, pady=4)

        # ── Glass Divider ──
        self.divider = ctk.CTkFrame(
            self.dock, width=1, height=28,
            corner_radius=0, fg_color=GLASS["divider"]
        )
        self.divider.pack(side="left", padx=5, pady=16)

        # ── Settings (Gear) ──
        self.btn_settings = ctk.CTkButton(
            self.dock, text="⚙", width=30, height=30,
            corner_radius=15, font=("SF Pro Display", 14),
            fg_color="transparent", hover_color=GLASS["glass_hover"],
            text_color=GLASS["text_dim"],
            command=self.show_settings_dialog
        )
        self.btn_settings.pack(side="left", padx=1, pady=4)

        # ── Close (X) ──
        self.btn_exit = ctk.CTkButton(
            self.dock, text="✕", width=28, height=28,
            corner_radius=14, font=("SF Pro Display", 12),
            fg_color="transparent", hover_color=GLASS["liquid_red"],
            text_color=GLASS["text_dim"],
            command=self.destroy
        )
        self.btn_exit.pack(side="left", padx=(1, 3), pady=4)

        # ── Glass Drag ──
        self.dock.bind("<Button-1>", self.start_move)
        self.dock.bind("<B1-Motion>", self.do_move)
        self.glass_shell.bind("<Button-1>", self.start_move)
        self.glass_shell.bind("<B1-Motion>", self.do_move)

        # ══════════════════════════════════════════════
        # ── PR Input Panel (Liquid Glass Expandable) ──
        # ══════════════════════════════════════════════
        self.pr_frame = ctk.CTkFrame(
            self.glass_shell, fg_color="transparent"
        )
        self.pr_frame.grid_columnconfigure(0, weight=2)   # Description
        self.pr_frame.grid_columnconfigure(1, weight=0)   # Checkbox
        self.pr_frame.grid_columnconfigure(2, weight=3)   # Account
        self.pr_frame.grid_columnconfigure(3, weight=1)   # Part No
        self.pr_frame.grid_columnconfigure(4, weight=0)   # Run
        self.pr_frame.grid_columnconfigure(5, weight=0)   # Close
        self.pr_frame.grid_rowconfigure(0, weight=1)

        # ── Glass Input Fields ──
        glass_entry = dict(
            corner_radius=12,
            border_width=1,
            border_color=GLASS["border_glass"],
            fg_color=GLASS["glass_elevated"],
            text_color=GLASS["text_bright"],
            font=("SF Pro Text", 13),
            height=34,
        )

        # Description
        self.desc_entry = ctk.CTkEntry(
            self.pr_frame, placeholder_text="PR Description",
            placeholder_text_color=GLASS["text_dim"],
            width=150, **glass_entry
        )
        self.desc_entry.grid(row=0, column=0, padx=(6, 3), pady=12, sticky="ew")

        # 단가 Checkbox (Glass toggle)
        self.unit_price_var = ctk.BooleanVar(value=False)
        self.unit_price_chk = ctk.CTkCheckBox(
            self.pr_frame, text="단가",
            variable=self.unit_price_var, width=48,
            font=("SF Pro Text", 11),
            text_color=GLASS["text_mid"],
            fg_color=GLASS["liquid_cyan"],
            hover_color=GLASS["glass_hover"],
            border_color=GLASS["border_glass"],
            corner_radius=6,
            checkbox_height=16, checkbox_width=16,
        )
        self.unit_price_chk.grid(row=0, column=1, padx=2, pady=12)

        # Account Code (Glass Dropdown)
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
            self.pr_frame, values=self.account_codes, width=150,
            corner_radius=12,
            border_width=1,
            border_color=GLASS["border_glass"],
            fg_color=GLASS["glass_elevated"],
            button_color=GLASS["glass_hover"],
            button_hover_color=GLASS["glass_pressed"],
            text_color=GLASS["text_bright"],
            font=("SF Pro Text", 11),
            dropdown_fg_color=GLASS["glass_surface"],
            dropdown_text_color=GLASS["text_bright"],
            dropdown_hover_color=GLASS["liquid_blue"],
            height=34,
        )
        self.account_combo.grid(row=0, column=2, padx=3, pady=12, sticky="ew")
        self.account_combo.set("0501030100/수선유지비")

        # Part No (Glass Input)
        self.part_entry = ctk.CTkEntry(
            self.pr_frame, placeholder_text="Part No",
            placeholder_text_color=GLASS["text_dim"],
            width=75, **glass_entry
        )
        self.part_entry.grid(row=0, column=3, padx=3, pady=12, sticky="ew")

        # ── Run Button (Liquid Green Glow) ──
        self.run_btn = ctk.CTkButton(
            self.pr_frame, text="▶", width=36, height=34,
            corner_radius=12,
            fg_color=GLASS["liquid_green"],
            hover_color=GLASS["liquid_green_dim"],
            text_color="white",
            font=("SF Pro Display", 15, "bold"),
            command=self.run_automation_thread
        )
        self.run_btn.grid(row=0, column=4, padx=3, pady=12)

        # ── Collapse Arrow (Glass) ──
        self.close_btn = ctk.CTkButton(
            self.pr_frame, text="‹", width=22, height=34,
            corner_radius=11,
            fg_color="transparent",
            hover_color=GLASS["glass_hover"],
            text_color=GLASS["text_dim"],
            font=("SF Pro Display", 17),
            command=self.toggle_pr_section
        )
        self.close_btn.grid(row=0, column=5, padx=(1, 5), pady=12)

        # Bind Escape to Hide
        self.bind("<Escape>", lambda e: self.hide_widget())

        # Internal State
        self.is_running = False
        self.pr_visible = False

    # ══════════════════════════════════════
    # ── Panel Toggle (Smooth Expand) ──
    # ══════════════════════════════════════
    def toggle_pr_section(self):
        if self.pr_visible:
            self.pr_frame.grid_forget()
            self.geometry(f"{WIDGET_WIDTH_MINI}x{WIDGET_HEIGHT}")
            self.btn_tools.configure(fg_color="transparent")
            self.pr_visible = False
        else:
            self.pr_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 2))
            self.geometry(f"{WIDGET_WIDTH_FULL}x{WIDGET_HEIGHT}")
            self.btn_tools.configure(fg_color=GLASS["liquid_blue"])
            self.pr_visible = True
            self.desc_entry.focus_set()

    # ══════════════════════════════════════
    # ── Settings Dialog (Frosted Glass) ──
    # ══════════════════════════════════════
    def show_settings_dialog(self):
        dialog = ctk.CTkToplevel(self)
        dialog.title("")
        dialog.overrideredirect(True)
        dialog.attributes('-topmost', True)
        dialog.attributes('-alpha', 0.94)
        dialog.configure(fg_color=GLASS["glass_bg"])

        # Position above widget
        wx, wy = self.winfo_x(), self.winfo_y()
        dw, dh = 280, 240
        dialog.geometry(f"{dw}x{dh}+{wx}+{wy - dh - 10}")

        # Glass container
        glass = ctk.CTkFrame(
            dialog, corner_radius=18,
            fg_color=GLASS["glass_surface"],
            border_width=1, border_color=GLASS["border_glass"]
        )
        glass.pack(fill="both", expand=True, padx=2, pady=2)

        # Title
        ctk.CTkLabel(
            glass, text="설정",
            font=("SF Pro Display", 15, "bold"),
            text_color=GLASS["text_bright"]
        ).pack(pady=(14, 6))

        # ── Startup Toggle ──
        startup_frame = ctk.CTkFrame(glass, fg_color="transparent")
        startup_frame.pack(fill="x", padx=18, pady=(0, 8))

        is_startup = self.is_in_startup()
        startup_var = ctk.BooleanVar(value=is_startup)

        def on_startup_toggle():
            self.toggle_startup(startup_var.get())

        ctk.CTkCheckBox(
            startup_frame, text="Windows 시작 시 자동 실행",
            variable=startup_var, command=on_startup_toggle,
            font=("SF Pro Text", 12),
            text_color=GLASS["text_bright"],
            fg_color=GLASS["liquid_green"],
            hover_color=GLASS["glass_hover"],
            border_color=GLASS["border_glass"],
            corner_radius=6,
            checkbox_height=16, checkbox_width=16
        ).pack(side="left")

        # ── Divider ──
        ctk.CTkFrame(glass, height=1, fg_color=GLASS["divider"]).pack(fill="x", padx=18, pady=4)

        # ── Password Section ──
        ctk.CTkLabel(
            glass, text="비밀번호 변경",
            font=("SF Pro Display", 12, "bold"),
            text_color=GLASS["text_mid"]
        ).pack(pady=(6, 4))

        pw_entry_cfg = dict(
            height=30, corner_radius=10,
            fg_color=GLASS["glass_elevated"],
            border_width=1, border_color=GLASS["border_glass"],
            font=("SF Pro Text", 12)
        )

        # Current password
        current_pw = menu_navigator.get_password()
        cur_frame = ctk.CTkFrame(glass, fg_color="transparent")
        cur_frame.pack(fill="x", padx=18, pady=2)
        ctk.CTkLabel(cur_frame, text="현재", width=36,
                     font=("SF Pro Text", 11),
                     text_color=GLASS["text_dim"]).pack(side="left")
        cur_entry = ctk.CTkEntry(cur_frame, text_color=GLASS["text_dim"], **pw_entry_cfg)
        cur_entry.pack(side="left", fill="x", expand=True, padx=(4, 0))
        cur_entry.insert(0, current_pw)
        cur_entry.configure(state="disabled")

        # New password
        new_frame = ctk.CTkFrame(glass, fg_color="transparent")
        new_frame.pack(fill="x", padx=18, pady=2)
        ctk.CTkLabel(new_frame, text="변경", width=36,
                     font=("SF Pro Text", 11),
                     text_color=GLASS["text_dim"]).pack(side="left")
        new_entry = ctk.CTkEntry(
            new_frame, text_color=GLASS["text_bright"],
            placeholder_text="새 비밀번호",
            placeholder_text_color=GLASS["text_dim"],
            **pw_entry_cfg
        )
        new_entry.pack(side="left", fill="x", expand=True, padx=(4, 0))

        # Buttons
        btn_frame = ctk.CTkFrame(glass, fg_color="transparent")
        btn_frame.pack(fill="x", padx=18, pady=(10, 12))

        def do_save():
            new_pw = new_entry.get().strip()
            if new_pw:
                menu_navigator.save_password(new_pw)
                print("Password updated successfully.")
            dialog.destroy()

        ctk.CTkButton(
            btn_frame, text="닫기", width=58, height=28,
            corner_radius=10, fg_color=GLASS["glass_elevated"],
            hover_color=GLASS["glass_hover"],
            text_color=GLASS["text_bright"],
            font=("SF Pro Text", 12),
            command=dialog.destroy
        ).pack(side="right", padx=(4, 0))

        ctk.CTkButton(
            btn_frame, text="저장", width=58, height=28,
            corner_radius=10, fg_color=GLASS["liquid_blue"],
            hover_color="#0060CC",
            text_color="white",
            font=("SF Pro Text", 12, "bold"),
            command=do_save
        ).pack(side="right")

        dialog.bind("<Escape>", lambda e: dialog.destroy())
        new_entry.bind("<Return>", lambda e: do_save())

    # ══════════════════════════
    # ── Startup Management ──
    # ══════════════════════════
    def get_startup_path(self):
        return os.path.join(
            os.getenv('APPDATA'),
            r'Microsoft\Windows\Start Menu\Programs\Startup',
            'PRMakerWidget.lnk'
        )

    def is_in_startup(self):
        return os.path.exists(self.get_startup_path())

    def toggle_startup(self, enable):
        shortcut_path = self.get_startup_path()
        if enable:
            try:
                import win32com.client
                shell = win32com.client.Dispatch("WScript.Shell")
                shortcut = shell.CreateShortCut(shortcut_path)
                if getattr(sys, 'frozen', False):
                    shortcut.Targetpath = sys.executable
                else:
                    shortcut.Targetpath = sys.executable
                    shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
                shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
                shortcut.save()
                print(f"Startup shortcut created.")
            except Exception as e:
                print(f"Failed to create startup shortcut: {e}")
        else:
            if os.path.exists(shortcut_path):
                try:
                    os.remove(shortcut_path)
                    print("Startup shortcut removed.")
                except Exception as e:
                    print(f"Failed to remove: {e}")

    # ══════════════════════════
    # ── Window Management ──
    # ══════════════════════════
    def center_window(self):
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        x = (sw - WIDGET_WIDTH_MINI) // 2
        y = sh - WIDGET_HEIGHT - 80
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
        dx = event.x - self._drag_x
        dy = event.y - self._drag_y
        self.geometry(f"+{self.winfo_x() + dx}+{self.winfo_y() + dy}")

    def show_at_cursor(self):
        try:
            mx, my = pyautogui.position()
            sw = self.winfo_screenwidth()
            sh = self.winfo_screenheight()
            w = WIDGET_WIDTH_FULL if self.pr_visible else WIDGET_WIDTH_MINI
            fx = min(mx, sw - w - 10)
            fy = my + 20 if my + WIDGET_HEIGHT + 20 < sh else my - WIDGET_HEIGHT - 10
            self.geometry(f"+{fx}+{fy}")
            self.bring_to_front()
            if self.pr_visible:
                self.desc_entry.focus_set()
        except:
            self.center_window()
            self.bring_to_front()

    def hide_widget(self):
        self.withdraw()

    # ══════════════════════════
    # ── Automation Runner ──
    # ══════════════════════════
    def run_automation_thread(self):
        if self.is_running:
            return
        desc = self.desc_entry.get()
        if not desc:
            # Flash border red
            self.desc_entry.configure(border_color=GLASS["liquid_red"])
            self.after(1500, lambda: self.desc_entry.configure(border_color=GLASS["border_glass"]))
            return
        account = self.account_combo.get()
        part_no = self.part_entry.get()
        self.run_btn.configure(state="disabled", fg_color=GLASS["glass_hover"])
        self.is_running = True
        threading.Thread(
            target=self._run_automation,
            args=(desc, account, part_no), daemon=True
        ).start()

    def _run_automation(self, desc, account, part_no):
        try:
            is_unit_price = self.unit_price_var.get()
            main.run_automation(desc, is_unit_price, account, part_no)
        except Exception as e:
            print(f"Automation Error: {e}")
        finally:
            self.is_running = False
            self.run_btn.configure(state="normal", fg_color=GLASS["liquid_green"])


# ══════════════════════════════
# ── Entry Point ──
# ══════════════════════════════
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
    print(" PR Maker Widget — Liquid Glass Edition")
    print("=" * 50)
    if PYNPUT_AVAILABLE:
        print(" [Ctrl + Shift + P] to show widget at cursor")
        threading.Thread(target=start_hotkey_listener, daemon=True).start()
    else:
        print(" [!] pynput not installed. Global hotkey disabled.")
    app.mainloop()
