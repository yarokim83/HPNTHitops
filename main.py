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
    import task_control as control
    import pr_form
    import roi_helpers
    import navigation
    description, part = pr_description.strip(), (part_no or '').strip()
    pr_form.validate(description, account_code, part)
    control.stage('Purchase Request 메뉴 여는 중')
    if not menu_navigator.click_pr_menu():
        return False
    control.stage("PR 목록 창 기다리는 중")
    hwnd = navigation.wait_window(roi_helpers.get_pr_window_rect)
    if not hwnd:
        return navigation.fail("PR: list window did not appear")
    if not menu_navigator.force_activate_window(hwnd):
        return navigation.fail("PR: cannot activate list window")
    return pr_form.fill(hwnd, description, is_unit_price, account_code, part)


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
