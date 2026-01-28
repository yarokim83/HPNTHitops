import subprocess
import os
import sys
import login_manager
import menu_navigator
import time
import pyautogui

def launch_hitops():
    """
    Launch Hitops3.exe from the specified path.
    """
    exe_path = r"C:\Program Files (x86)\Hyundai-UNI\HITOPSIII\Hitops3.exe"
    
    if not os.path.exists(exe_path):
        print(f"Error: Executable not found at {exe_path}")
        return

    import tkinter as tk
    from tkinter import simpledialog, messagebox
    from tkinter import ttk

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
        
        # Account Codes (Transcribed from screenshot)
        account_codes = [
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
        
        # UI Elements
        tk.Label(dialog, text="Enter PR Description (Title):").pack(pady=5)
        entry = tk.Entry(dialog, textvariable=desc_var, width=60)
        entry.pack(pady=5)
        entry.focus_set()
        
        tk.Checkbutton(dialog, text="Unit Price Contract (단가계약)", variable=unit_price_var).pack(pady=5)
        
        tk.Label(dialog, text="Select Account Code:").pack(pady=5)
        code_combo = ttk.Combobox(dialog, textvariable=account_code_var, values=account_codes, width=57)
        code_combo.pack(pady=5)
        if account_codes:
            code_combo.current(0)
        
        result = {"description": None, "is_unit_price": False, "account_code": None}
        
        def on_ok():
            result["description"] = desc_var.get()
            result["is_unit_price"] = unit_price_var.get()
            result["account_code"] = account_code_var.get()
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

        if not pr_description:
            print("No description entered. Exiting.")
            return

        print(f"Description: {pr_description}, Unit Price: {is_unit_price}, Account Code: {account_code}")

        # 2. Launch Hitops3
        working_dir = os.path.dirname(exe_path)
        print(f"Launching {exe_path}...")
        subprocess.Popen(exe_path, cwd=working_dir)
        
        # 3. Wait for Login Screen
        time.sleep(5) 
        
        # 4. Perform Login
        password = "fdjk213!@"
        if login_manager.perform_login(password):
            print("Login successful (or already logged in). Proceeding...")
        else:
            print("Login failed or timed out.")
            return

        # 5. Navigate Menu
        menu_navigator.navigate_to_mr()
        time.sleep(2) # Wait for new screen/menu to load
        menu_navigator.click_inventory()
        time.sleep(2)
        menu_navigator.click_purchase_request()
        
        # 6. Click Add Button
        time.sleep(2)
        menu_navigator.click_add_button()

        # --- Enter PR Description ---
        time.sleep(2) # Wait for form to open
        menu_navigator.enter_pr_description(pr_description)
        time.sleep(2) # Increased delay to ensure Need By field is ready
        menu_navigator.update_need_by_date()
        time.sleep(1)
        menu_navigator.set_unit_price_contract(is_unit_price)
        time.sleep(1)
        menu_navigator.set_account_code(account_code)
        
    except Exception as e:
        print(f"Failed to launch application: {e}")

if __name__ == "__main__":
    launch_hitops()
