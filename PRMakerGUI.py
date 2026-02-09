import customtkinter as ctk
import tkinter as tk
import threading
import sys
import os
import time
import main

# Compact Configuration
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        try:
            self.widget.configure(state="normal")
            self.widget.insert("end", str, (self.tag,))
            self.widget.see("end")
            self.widget.configure(state="disabled")
        except:
            pass
        
    def flush(self):
        pass

class PRMakerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Setup Window
        self.title("PR Maker V2.1")
        self.geometry("500x420")
        self.resizable(False, False)

        # Account Codes Data
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

        self.create_widgets()

    def create_widgets(self):
        # 1. Main Input Frame (Compact)
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="x", padx=20, pady=(20, 10))

        # Description (No Label, just bold placeholder)
        self.desc_entry = ctk.CTkEntry(self.main_frame, placeholder_text="PR Description (Title) - Essential", height=40, font=("Arial", 14))
        self.desc_entry.pack(fill="x", pady=(0, 12))

        # Account Code (Label + Combo)
        code_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        code_frame.pack(fill="x", pady=(0, 12))
        
        ctk.CTkLabel(code_frame, text="Account:", width=60, anchor="w", font=("Arial", 12)).pack(side="left")
        self.code_combo = ctk.CTkComboBox(code_frame, values=self.account_codes, height=32, font=("Arial", 12))
        self.code_combo.pack(side="left", fill="x", expand=True, padx=(5, 0))
        self.code_combo.set(self.account_codes[0])

        # Part No & Unit Price (Side by Side)
        option_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        option_frame.pack(fill="x", pady=(0, 5))
        
        # Part No
        self.part_entry = ctk.CTkEntry(option_frame, placeholder_text="Part No (Opt)", width=140, height=32)
        self.part_entry.pack(side="left", padx=(0, 15))
        
        # Unit Price Check (Switch is cleaner)
        self.unit_price_var = tk.BooleanVar(value=False)
        self.unit_price_check = ctk.CTkSwitch(option_frame, text="Unit Price", variable=self.unit_price_var, 
                                            font=("Arial", 12))
        self.unit_price_check.pack(side="left")

        # 2. Control Buttons
        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=20, pady=5)

        self.start_btn = ctk.CTkButton(self.btn_frame, text="▶ EXECUTE", command=self.start_automation, 
                                     fg_color="#106EBE", hover_color="#005A9E", height=45, font=("Arial", 13, "bold"))
        self.start_btn.pack(side="left", expand=True, fill="x", padx=(0, 10))

        self.stop_btn = ctk.CTkButton(self.btn_frame, text="■ STOP", command=self.stop_automation, 
                                    fg_color="#D93025", hover_color="#B31412", state="disabled", height=45, width=90, font=("Arial", 12, "bold"))
        self.stop_btn.pack(side="right")

        # 3. Compact Log
        self.log_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.log_frame.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        
        # Tiny header for log
        log_header = ctk.CTkFrame(self.log_frame, height=20, fg_color="transparent")
        log_header.pack(fill="x")
        ctk.CTkLabel(log_header, text="STATUS LOG", font=("Arial", 10, "bold"), text_color="gray").pack(side="left")
        
        self.log_text = ctk.CTkTextbox(self.log_frame, height=100, font=("Consolas", 11), state="disabled", fg_color="#1e1e1e")
        self.log_text.pack(fill="both", expand=True)

        # Redirect Stdout
        sys.stdout = TextRedirector(self.log_text, "stdout")
        sys.stderr = TextRedirector(self.log_text, "stderr")

    def start_automation(self):
        description = self.desc_entry.get()
        if not description:
            print("Error: PR Description is required!")
            return

        is_unit_price = self.unit_price_var.get()
        account_code = self.code_combo.get()
        part_no = self.part_entry.get()

        self.start_btn.configure(state="disabled", text="RUNNING...")
        self.stop_btn.configure(state="normal")
        self.desc_entry.configure(state="disabled") # Lock inputs

        print(f"Starting Automation...\nDesc: {description}\nCode: {account_code}\nPart: {part_no}")

        # Run in Thread
        self.thread = threading.Thread(target=self.run_logic, args=(description, is_unit_price, account_code, part_no))
        self.thread.daemon = True
        self.thread.start()

    def run_logic(self, desc, unit, code, part):
        try:
            main.run_automation(desc, unit, code, part)
            print("\n[SUCCESS] Automation Finished.")
        except Exception as e:
            print(f"\n[ERROR] Automation Failed: {e}")
        finally:
            # Schedule UI reset on main thread
            self.after(100, self.reset_ui)

    def reset_ui(self):
        self.start_btn.configure(state="normal", text="▶ EXECUTE")
        self.stop_btn.configure(state="disabled")
        self.desc_entry.configure(state="normal")
        print("Ready.")

    def stop_automation(self):
        print("\n[STOP] Force Stop requested. Please restart the app if stuck.")
        # In this simple version, we can't easily kill the thread safely without shared flags.
        # But we can unlock the UI.
        self.reset_ui()

if __name__ == "__main__":
    app = PRMakerApp()
    app.mainloop()
