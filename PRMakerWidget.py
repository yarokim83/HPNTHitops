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

# Widget Configuration
WIDGET_WIDTH_FULL = 1200
WIDGET_WIDTH_MINI = 520
WIDGET_HEIGHT = 160
ALPHA_VALUE = 0.95

class PRMakerWidget(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Assets
        self.assets_dir = os.path.join(os.path.dirname(__file__), 'assets')
        
        # Initialize OCR Engine (Essential for Vessel detection)
        threading.Thread(target=ocr_helpers.init_tesseract, daemon=True).start()

        # Window Setup
        self.title("PR Maker Widget")
        
        # Start Collapsed (Mini Mode)
        self.geometry(f"{WIDGET_WIDTH_MINI}x{WIDGET_HEIGHT}")
        self.overrideredirect(True) # Frameless
        self.attributes('-alpha', ALPHA_VALUE)
        self.resizable(False, False)
        
        # Center initially
        self.center_window()
        
        # Layout: Sidebar (0) | Divider (1) | PR Input Section (2)
        self.grid_columnconfigure(0, weight=0) # Sidebar
        self.grid_columnconfigure(1, weight=0) # Divider
        self.grid_columnconfigure(2, weight=1) # Inputs
        self.grid_rowconfigure(0, weight=1)
        
        # --- Sidebar (Mode Icons) ---
        self.sidebar = ctk.CTkFrame(self, corner_radius=10, fg_color="#2B2B2B")
        self.sidebar.grid(row=0, column=0, padx=5, pady=5, sticky="ns")
        
        # Load Images
        try:
            from PIL import Image
            self.img_tools = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "tools_icon.png")), size=(80, 80))
            self.img_mc = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "mc_icon.png")), size=(80, 80))
            self.img_rcc = ctk.CTkImage(light_image=Image.open(os.path.join(self.assets_dir, "rcc_icon.png")), size=(80, 80))
        except Exception as e:
            print(f"Failed to load icons: {e}")
            self.img_tools = self.img_mc = self.img_rcc = None

        # Mode Buttons
        self.btn_tools = ctk.CTkButton(self.sidebar, image=self.img_tools, text="", width=100, height=100, 
                                      fg_color="transparent", hover_color="#555555", command=self.toggle_pr_section)
        self.btn_tools.pack(side="left", padx=8)
        
        self.btn_mc = ctk.CTkButton(self.sidebar, image=self.img_mc, text="", width=100, height=100, 
                                   fg_color="transparent", hover_color="#444444", command=lambda: threading.Thread(target=menu_navigator.run_mc_sequence, daemon=True).start())
        self.btn_mc.pack(side="left", padx=8)
        
        self.btn_rcc = ctk.CTkButton(self.sidebar, image=self.img_rcc, text="", width=100, height=100, 
                                    fg_color="transparent", hover_color="#444444", command=lambda: threading.Thread(target=menu_navigator.click_rcc_menu, daemon=True).start())
        self.btn_rcc.pack(side="left", padx=8)

        # Exit Button
        self.btn_exit = ctk.CTkButton(self.sidebar, text="✕", width=100, height=100, font=("Arial", 28, "bold"),
                                     fg_color="#661111", hover_color="#AA2222", text_color="white", command=self.destroy)
        self.btn_exit.pack(side="left", padx=8)

        # Drag Handle
        self.drag_handle = ctk.CTkFrame(self, width=10, corner_radius=0, fg_color="#333333")
        self.drag_handle.grid(row=0, column=1, sticky="ns", padx=(2, 2))
        self.drag_handle.bind("<Button-1>", self.start_move)
        self.drag_handle.bind("<B1-Motion>", self.do_move)
        
        # --- PR Input Section (Hidden initially) ---
        self.pr_frame = ctk.CTkFrame(self, fg_color="transparent")
        # Grid inside pr_frame
        self.pr_frame.grid_columnconfigure((0, 2), weight=1)
        self.pr_frame.grid_rowconfigure(0, weight=1)

        # 2. Description Input
        self.desc_entry = ctk.CTkEntry(self.pr_frame, placeholder_text="PR Description", width=200)
        self.desc_entry.grid(row=0, column=0, padx=5, pady=10, sticky="ew")
        
        # 3. Checkbox (Unit Price)
        self.unit_price_var = ctk.BooleanVar(value=False)
        self.unit_price_chk = ctk.CTkCheckBox(self.pr_frame, text="단가", variable=self.unit_price_var, width=60)
        self.unit_price_chk.grid(row=0, column=1, padx=5, pady=10)

        # 4. Account Code Combo
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
        self.account_combo = ctk.CTkComboBox(self.pr_frame, values=self.account_codes, width=180)
        self.account_combo.grid(row=0, column=2, padx=5, pady=10, sticky="ew")
        self.account_combo.set("0501030100/수선유지비")
        
        # 5. Part No Input
        self.part_entry = ctk.CTkEntry(self.pr_frame, placeholder_text="Part No (Opt)", width=100)
        self.part_entry.grid(row=0, column=3, padx=5, pady=10, sticky="ew")
        
        # 6. Run Button
        self.run_btn = ctk.CTkButton(self.pr_frame, text="▶", width=40, fg_color="green", command=self.run_automation_thread)
        self.run_btn.grid(row=0, column=4, padx=5, pady=10)
        
        # 7. Close Button
        self.close_btn = ctk.CTkButton(self.pr_frame, text="X", width=30, fg_color="#444444", hover_color="#882222", command=self.hide_widget)
        self.close_btn.grid(row=0, column=5, padx=(5, 10), pady=10)
        
        # Bind Escape to Hide
        self.bind("<Escape>", lambda e: self.hide_widget())
        
        # Internal State
        self.is_running = False
        self.pr_visible = False # Start collapsed by default

    def toggle_pr_section(self):
        if self.pr_visible:
            self.pr_frame.grid_forget()
            self.geometry(f"{WIDGET_WIDTH_MINI}x{WIDGET_HEIGHT}")
            self.btn_tools.configure(fg_color="transparent")
            self.pr_visible = False
        else:
            self.pr_frame.grid(row=0, column=2, sticky="nsew", padx=5)
            self.geometry(f"{WIDGET_WIDTH_FULL}x{WIDGET_HEIGHT}")
            self.btn_tools.configure(fg_color="#3D3D3D")
            self.pr_visible = True

    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - WIDGET_WIDTH_FULL) // 2
        y = (screen_height - WIDGET_HEIGHT) // 2
        self.geometry(f"+{x}+{y}")

    def bring_to_front(self):
        self.deiconify()
        self.attributes('-topmost', True)
        self.lift()
        self.focus_force()
        self.after(200, lambda: self.attributes('-topmost', False))

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def do_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
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
        if self.is_running: return
        desc = self.desc_entry.get()
        if not desc:
            print("Description is required!")
            return
        account = self.account_combo.get()
        part_no = self.part_entry.get()
        self.run_btn.configure(state="disabled")
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
            self.run_btn.configure(state="normal")

app = None

def on_activate():
    if app:
        app.after(0, app.show_at_cursor)

def start_hotkey_listener():
    if not PYNPUT_AVAILABLE: return
    with keyboard.GlobalHotKeys({'<ctrl>+<shift>+p': on_activate}) as h:
        h.join()

if __name__ == "__main__":
    app = PRMakerWidget()
    print("="*50)
    print(" PR Maker Widget Mode Enhanced")
    print("="*50)
    if PYNPUT_AVAILABLE:
        print(" [Ctrl + Shift + P] to show widget at mouse cursor")
        threading.Thread(target=start_hotkey_listener, daemon=True).start()
    else:
        print(" [!] pynput not installed. Global hotkey disabled.")
    app.mainloop()
