import pyautogui
import pyperclip
import time
import os
import sys
from datetime import datetime
from dateutil.relativedelta import relativedelta

def fill_pr_form_via_tabs(description, is_unit_price=False, account_code=None, part_no=None):
    """
    Fill the PR form using keyboard Tab navigation.
    
    After clicking 'Add', the cursor is in the Description field.
    Tab order in HiTOPS PR Form (MNR035 Purchase Requisition Detail):
      Description -> [View btn] -> Need By date -> Urgent dropdown -> 단가계약 dropdown -> ...
    
    This approach is monitor/DPI-independent.
    """
    
    # --- 1. Description (cursor already here after Add) ---
    print(f"Entering Description: {description}")
    time.sleep(0.5)
    # Use clipboard for Korean text support
    pyperclip.copy(description)
    pyautogui.hotkey('ctrl', 'v')
    time.sleep(0.3)
    
    # --- 2. Need By Date ---
    # Tab to Need By field: Description -> [View button] -> Need By
    # We need to skip the View button
    print("Tabbing to Need By field...")
    pyautogui.press('tab')  # Skip View button
    time.sleep(0.2)
    pyautogui.press('tab')  # Land on Need By
    time.sleep(0.2)
    
    # Calculate date = today + 1 month
    new_date = (datetime.now() + relativedelta(months=1)).strftime('%Y-%m-%d')
    print(f"Entering Need By Date: {new_date}")
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(0.1)
    pyautogui.write(new_date, interval=0.02)
    time.sleep(0.3)
    
    # --- 3. Unit Price (단가계약) ---
    # Tab from Need By -> Urgent -> 단가계약
    print("Tabbing to 단가계약 field...")
    pyautogui.press('tab')  # Urgent dropdown
    time.sleep(0.2)
    pyautogui.press('tab')  # 단가계약 dropdown
    time.sleep(0.2)
    
    if is_unit_price:
        print("Setting 단가계약 to Y...")
        # Open dropdown and select Y
        pyautogui.press('space')  # or click to open
        time.sleep(0.2)
        pyautogui.press('down')  # Move to Y
        time.sleep(0.1)
        pyautogui.press('enter')
        time.sleep(0.2)
    
    # --- 4. Account Code ---
    # The Account Code field is further down in "Purchase Properties" section
    # We'll skip for now and navigate via more tabs or image search
    # Tab from 단가계약 -> ... -> Account Code
    # This needs to be verified with the actual form
    
    print(f"Setting Account Code: {account_code}")
    # Skip fields: Etc.Process checkbox -> Cost Year -> Account Code
    pyautogui.press('tab')  # Etc. Process
    time.sleep(0.1)
    pyautogui.press('tab')  # Cost Year
    time.sleep(0.1) 
    pyautogui.press('tab')  # Account Code dropdown
    time.sleep(0.2)
    
    if account_code:
        # Extract just the code number (before the /)
        code_num = account_code.split('/')[0] if '/' in account_code else account_code
        # Type the code number to search in dropdown
        pyautogui.hotkey('ctrl', 'a')
        time.sleep(0.1)
        pyautogui.write(code_num, interval=0.02)
        time.sleep(0.3)
    
    # --- 5. Part No ---
    if part_no and len(str(part_no).strip()) > 3:
        # Navigate to Part No section (bottom of form)
        # Tab through remaining Purchase Properties fields
        print(f"Entering Part No: {part_no}")
        # Type of Purchasing -> Kind -> Type of Purchase -> Included in Plan -> Part No.
        for _ in range(4):
            pyautogui.press('tab')
            time.sleep(0.1)
        pyautogui.press('tab')  # Part No field
        time.sleep(0.2)
        pyperclip.copy(str(part_no))
        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.3)
        print("Part No entered.")
    
    print("\n" + "="*60)
    print("✓ PR Form filled successfully via Tab navigation!")
    print("="*60)

if __name__ == "__main__":
    fill_pr_form_via_tabs("Test Description Tab Nav", is_unit_price=False, account_code="0501030100", part_no="12345")
