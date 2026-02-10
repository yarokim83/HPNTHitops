import customtkinter as ctk
import pyautogui
import threading
import sys
import os
import time
import main

# Check for pynput
try:
    from pynput import keyboard
    PYNPUT_AVAILABLE = True
except ImportError:
    PYNPUT_AVAILABLE = False
    print("Warning: 'pynput' library not found. Global hotkey will not work.")
    print("Please install it: pip install pynput")

# Widget Configuration
WIDGET_WIDTH = 620
WIDGET_HEIGHT = 80
ALPHA_VALUE = 0.95

class PRMakerWidget(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("PR Maker Widget")
        self.geometry(f"{WIDGET_WIDTH}x{WIDGET_HEIGHT}")
        self.overrideredirect(True) # Frameless
        # self.attributes('-topmost', True) # Removed permanent Topmost
        self.attributes('-alpha', ALPHA_VALUE)
        self.resizable(False, False)
        
        # Center initially (or hide)
        self.center_window()
        
        # Grid Layout
        self.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # --- UI Components ---
        
        # 1. Drag Handle (Left side)
        self.drag_handle = ctk.CTkFrame(self, width=20, corner_radius=0, fg_color="#333333")
        self.drag_handle.grid(row=0, column=0, sticky="ns", padx=(0, 5))
        self.drag_handle.bind("<Button-1>", self.start_move)
        self.drag_handle.bind("<B1-Motion>", self.do_move)
        
        # 2. Description Input
        self.desc_entry = ctk.CTkEntry(self, placeholder_text="PR Description", width=200)
        self.desc_entry.grid(row=0, column=1, padx=5, pady=10, sticky="ew")
        
        # 3. Checkbox (Unit Price)
        self.unit_price_var = ctk.BooleanVar(value=False)
        self.unit_price_chk = ctk.CTkCheckBox(self, text="단가", variable=self.unit_price_var, width=60)
        self.unit_price_chk.grid(row=0, column=2, padx=5, pady=10)

        # 4. Account Code Combo
        self.account_codes = [
            "0501030000/수선유지비",
            "0501030100/수선유지비",
            "0501030101/장비 자재비-QC",
            "0501030102/장비 자재비-ATC",
            "0501030103/장비 자재비-RS",
            "0501030104/장비 자재비-YT",
            "0501030105/장비 자재비-YC",
            "0501030106/장비 자재비-FL",
            "0501030107/장비 자재비-기타",
            "0501030108/수선유지비-외주수리-QC",
            "0501030109/수선유지비-외주수리-ATC",
            "0501030110/수선유지비-외주수리-RS",
            "0501030111/수선유지비-외주수리-YT",
            "0501030112/수선유지비-외주수리-YC",
            "0501030113/수선유지비-외주수리-FL",
            "0501030114/수선유지비-외주수리-기타",
            "0501030115/시설물-야드시설물(자재)",
            "0501030116/수선유지비-시설물-CFS시설물",
            "0501030117/시설물-전기시설물(자재)",
            "0501030118/시설물-외주수리",
            "0501030119/수선유지비_작업공구-야드공구",
            "0501030120/수선유지비_작업공구-정비공구",
            "0501030121/수선유지비_작업공구-CFS공구",
            "0501030122/수선유지비_작업공구-안전공구",
            "0501030123/수선유지비_작업공구-기타공구",
            "0501030124/수선유지비_작업소모품-야드소모품",
            "0501030125/작업소모품-정비소모품/공구",
            "0501030126/수선유지비_작업소모품-CFS소모품",
            "0501030127/수선유지비_작업소모품-안전소모품",
            "0501030128/수선유지비_작업소모품-기타소모품",
            "0501030129/수선유지비-CNTR",
            "0501030130/수선유지비-기타 (사고변상금등)",
            "0501030131/장비자재비-ECH",
            "0501030132/수선유지비-외주수리-ECH",
            "0501040106/동력비-윤활유"
        ]
        self.account_combo = ctk.CTkComboBox(self, values=self.account_codes, width=180)
        self.account_combo.grid(row=0, column=3, padx=5, pady=10, sticky="ew")
        self.account_combo.set("0501030100/수선유지비") # Default
        
        # 5. Part No Input
        self.part_entry = ctk.CTkEntry(self, placeholder_text="Part No (Opt)", width=100)
        self.part_entry.grid(row=0, column=4, padx=5, pady=10, sticky="ew")
        
        # 6. Run Button
        self.run_btn = ctk.CTkButton(self, text="▶", width=40, fg_color="green", command=self.run_automation_thread)
        self.run_btn.grid(row=0, column=5, padx=5, pady=10)
        
        # 7. Close Button
        self.close_btn = ctk.CTkButton(self, text="X", width=30, fg_color="#444444", hover_color="#882222", command=self.hide_widget)
        self.close_btn.grid(row=0, column=6, padx=(5, 10), pady=10)
        
        # Bind Escape to Hide
        self.bind("<Escape>", lambda e: self.hide_widget())
        
        # Internal State
        self.is_running = False

    def center_window(self):
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - WIDGET_WIDTH) // 2
        y = (screen_height - WIDGET_HEIGHT) // 2
        self.geometry(f"+{x}+{y}")

    def start_move(self, event):
        self.x = event.x
        self.y = event.y

    def do_move(self, event):
        deltax = event.x - self.x
        deltay = event.y - self.y
        x = self.winfo_x() + deltax
        y = self.winfo_y() + deltay
        self.geometry(f"+{x}+{y}")
        
    def bring_to_front(self):
        self.deiconify()
        self.attributes('-topmost', True)
        self.lift()
        self.focus_force()
        # Disable topmost after a short delay to allow other windows to cover it
        self.after(200, lambda: self.attributes('-topmost', False))

    def show_at_cursor(self):
        try:
            # Get Mouse Position
            mouse_x, mouse_y = pyautogui.position()
            
            # Adjust to keep within screen
            screen_w = self.winfo_screenwidth()
            screen_h = self.winfo_screenheight()
            
            final_x = mouse_x
            final_y = mouse_y + 20 # Below cursor
            
            if final_x + WIDGET_WIDTH > screen_w:
                final_x = screen_w - WIDGET_WIDTH - 10
            if final_y + WIDGET_HEIGHT > screen_h:
                final_y = mouse_y - WIDGET_HEIGHT - 10
                
            self.geometry(f"+{final_x}+{final_y}")
            self.bring_to_front()
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
        
        # Disable UI
        self.run_btn.configure(state="disabled")
        self.is_running = True
        
        # Run in thread
        threading.Thread(target=self._run_automation, args=(desc, account, part_no), daemon=True).start()
        
    def _run_automation(self, desc, account, part_no):
        try:
            # Call Main Logic
            is_unit_price = self.unit_price_var.get()
            main.run_automation(desc, is_unit_price, account, part_no)
        except Exception as e:
            print(f"Automation Error: {e}")
        finally:
            self.is_running = False
            self.run_btn.configure(state="normal")
            
            # Auto-hide after success?
            # self.after(1000, self.hide_widget)


# Global variable for Hotkey Thread
app = None

def on_activate():
    if app:
        # Schedule show_at_cursor in Main Thread
        app.after(0, app.show_at_cursor)

def start_hotkey_listener():
    if not PYNPUT_AVAILABLE:
        return
        
    # Hotkey: Ctrl + Shift + P
    with keyboard.GlobalHotKeys({'<ctrl>+<shift>+p': on_activate}) as h:
        h.join()

if __name__ == "__main__":
    app = PRMakerWidget()
    
    # Instructions
    print("="*50)
    print(" PR Maker Widget Mode")
    print("="*50)
    if PYNPUT_AVAILABLE:
        print(" [Ctrl + Shift + P] to show widget at mouse cursor")
        # Start Hotkey Listener in separate thread
        t = threading.Thread(target=start_hotkey_listener, daemon=True)
        t.start()
    else:
        print(" [!] pynput not installed. Global hotkey disabled.")
        print("     Run 'pip install pynput' to enable hotkeys.")
        print("     Widget will start visible.")
        
    app.mainloop()
