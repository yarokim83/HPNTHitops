import customtkinter as ctk
import tkinter as tk
import threading
import sys
import os
import time

# Import Automation Logic
import main

# Configuration
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.see("end")
        self.widget.configure(state="disabled")
        
    def flush(self):
        pass

class PRMakerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("PR Maker V2.0 (Modern UI)")
        self.geometry("700x650")
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
        # Header
        self.header_frame = ctk.CTkFrame(self)
        self.header_frame.pack(fill="x", padx=20, pady=20)
        
        self.header_label = ctk.CTkLabel(self.header_frame, text="HI-TOPS PR Automation", font=ctk.CTkFont(size=20, weight="bold"))
        self.header_label.pack(pady=10)

        # Inputs Frame
        self.input_frame = ctk.CTkFrame(self)
        self.input_frame.pack(fill="x", padx=20, pady=10)

        # 1. Description
        self.desc_label = ctk.CTkLabel(self.input_frame, text="PR Description (Title):")
        self.desc_label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.desc_entry = ctk.CTkEntry(self.input_frame, width=400, placeholder_text="Enter PR Description...")
        self.desc_entry.grid(row=0, column=1, padx=10, pady=10)

        # 2. Account Code
        self.code_label = ctk.CTkLabel(self.input_frame, text="Account Code:")
        self.code_label.grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.code_combo = ctk.CTkComboBox(self.input_frame, values=self.account_codes, width=400)
        self.code_combo.grid(row=1, column=1, padx=10, pady=10)
        self.code_combo.set(self.account_codes[0])

        # 3. Part No (Optional)
        self.part_label = ctk.CTkLabel(self.input_frame, text="Part No (Optional):")
        self.part_label.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        self.part_entry = ctk.CTkEntry(self.input_frame, width=400, placeholder_text="Enter Part No (min 4 chars)...")
        self.part_entry.grid(row=2, column=1, padx=10, pady=10)

        # 4. Unit Price Checkbox
        self.unit_price_var = tk.BooleanVar(value=False)
        self.unit_price_check = ctk.CTkCheckBox(self.input_frame, text="Unit Price Contract (단가계약)", variable=self.unit_price_var)
        self.unit_price_check.grid(row=3, column=1, padx=10, pady=10, sticky="w")

        # Buttons Frame
        self.btn_frame = ctk.CTkFrame(self)
        self.btn_frame.pack(fill="x", padx=20, pady=10)

        self.start_btn = ctk.CTkButton(self.btn_frame, text="START AUTOMATION", command=self.start_automation, fg_color="green", height=40)
        self.start_btn.pack(side="left", expand=True, fill="x", padx=10, pady=10)

        self.stop_btn = ctk.CTkButton(self.btn_frame, text="STOP", command=self.stop_automation, fg_color="red", state="disabled", height=40)
        self.stop_btn.pack(side="right", expand=True, fill="x", padx=10, pady=10)

        # Console Output Area
        self.log_label = ctk.CTkLabel(self, text="Execution Log:")
        self.log_label.pack(anchor="w", padx=20, pady=(10, 0))
        
        self.log_text = ctk.CTkTextbox(self, width=660, height=200, state="disabled")
        self.log_text.pack(padx=20, pady=10)

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

        self.start_btn.configure(state="disabled")
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
            self.after(100, self.reset_ui) # Validate UI update on main thread

    def reset_ui(self):
        self.start_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")
        self.desc_entry.configure(state="normal")
        print("Ready for next task.")

    def stop_automation(self):
        print("\n[STOP] Force Stop requested. (Note: Logic thread might not stop immediately)")
        # In a real app, we'd use a stop_event flag in main code to check periodically.
        # For now, we just reset UI, though background thread continues until next check.
        self.reset_ui()
        # TODO: Implement graceful stop in main.py loop

if __name__ == "__main__":
    app = PRMakerApp()
    app.mainloop()
