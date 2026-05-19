import subprocess
import os
import sys
import login_manager
import menu_navigator
import time
import pyautogui
import win32api
import win32con
import tkinter as tk
from tkinter import simpledialog, messagebox
from tkinter import ttk
from account_codes import ACCOUNT_CODES

def run_automation(pr_description, is_unit_price, account_code, part_no=None):
    """
    Core automation logic, decoupled from UI.
    Executing this function runs the full PR creation flow.
    """
    exe_path = r"C:\Program Files (x86)\Hyundai-UNI\HITOPSIII\Hitops3.exe"
    if not os.path.exists(exe_path):
        print(f"Error: Executable not found at {exe_path}")
        return

    try:
        # 1. Common Launch & Login & Maximize (Shared with M&C flow)
        if not menu_navigator.ensure_app_ready():
            print("App initialization failed. Aborting PR automation.")
            return

        # 2. Smart Menu Navigation (Parallel/Event-Driven)
        print("Executing Smart Navigation...")
        if not menu_navigator.smart_navigate_to_pr():
             print("Smart Navigation failed or timed out.")
             return
             
        # Form is now presumably open. Appending verify logic or wait.
        time.sleep(0.5) 
        
        # 6. Click Add Button
        time.sleep(0.1)
        menu_navigator.click_add_button()

        # --- Enter PR Description ---
        time.sleep(0.2) # Wait for form to open
        menu_navigator.enter_pr_description(pr_description)
        time.sleep(0.1)
        menu_navigator.update_need_by_date()
        time.sleep(0.1)
        menu_navigator.set_unit_price_contract(is_unit_price)
        time.sleep(0.1)
        menu_navigator.set_account_code(account_code)
        
        # New: Enter Part No (if provided and valid)
        if part_no and len(str(part_no).strip()) > 3:
            time.sleep(0.1)
            menu_navigator.enter_part_no(part_no)
        else:
            print(f"Skipping Part No input: '{part_no}' is too short or potentially unsafe.")
        
        # 7. Program complete
        print("\n" + "="*60)
        print("✓ Part No entry complete!")
        print("✓ Automation stopped as requested after Part No.")
        print("="*60)
        return

    except Exception as e:
        print(f"Failed to run automation: {e}")
        raise e

def launch_hitops():
    """
    Legacy Entry Point: User Input via Dialog -> calling run_automation
    """
    # Custom Dialog to get Description + Checkbox + Account Code
    def get_user_input():
        root = tk.Tk()
        root.withdraw() # Hide main window
        root.attributes('-topmost', True) # Keep on top
        
        dialog = tk.Toplevel(root)
        dialog.title("PR Maker Input")
        dialog.geometry("500x300")
        dialog.attributes('-topmost', True)
        
        # Variables
        desc_var = tk.StringVar()
        unit_price_var = tk.BooleanVar()
        account_code_var = tk.StringVar()
        
        # Account Codes — single source of truth
        account_codes = ACCOUNT_CODES
        
        # UI Elements
        tk.Label(dialog, text="Enter PR Description (Title):").pack(pady=5)
        entry = tk.Entry(dialog, textvariable=desc_var, width=60)
        entry.pack(pady=5)
        entry.focus_set()
        
        tk.Checkbutton(dialog, text="Unit Price Contract (단가계약)", variable=unit_price_var).pack(pady=5)
        
        # Part No Input
        tk.Label(dialog, text="Enter Part No (Optional):").pack(pady=5)
        part_no_var = tk.StringVar()
        entry_part_no = tk.Entry(dialog, textvariable=part_no_var, width=60)
        entry_part_no.pack(pady=5)
        
        tk.Label(dialog, text="Select Account Code:").pack(pady=5)
        code_combo = ttk.Combobox(dialog, textvariable=account_code_var, values=account_codes, width=57)
        code_combo.pack(pady=5)
        if account_codes:
            code_combo.current(0)
        
        result = {"description": None, "is_unit_price": False, "account_code": None, "part_no": None}
        
        def on_ok():
            result["description"] = desc_var.get()
            result["is_unit_price"] = unit_price_var.get()
            result["account_code"] = account_code_var.get()
            result["part_no"] = part_no_var.get()
            dialog.destroy()
            root.destroy()
            
        def on_cancel():
            dialog.destroy()
            root.destroy()
            
        tk.Button(dialog, text="OK", command=on_ok, width=10).pack(side=tk.LEFT, padx=50, pady=20)
        tk.Button(dialog, text="Cancel", command=on_cancel, width=10).pack(side=tk.RIGHT, padx=50, pady=20)
        
        root.wait_window(dialog)
        return result

    try:
        # 1. Get Input from user
        user_input = get_user_input()
        pr_description = user_input["description"]
        is_unit_price = user_input["is_unit_price"]
        account_code = user_input["account_code"]
        part_no = user_input["part_no"]

        if not pr_description:
            print("No description entered. Exiting.")
            return
            
        # 2. Run Automation
        run_automation(pr_description, is_unit_price, account_code, part_no)

    except Exception as e:
        print(f"Failed to launch application: {e}")

if __name__ == "__main__":
    launch_hitops()
