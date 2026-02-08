import pywinauto
from pywinauto import Desktop
import sys

def inspect_uitree():
    print("Inspecting UI Controls for 'Hitops'...")
    
    # List all windows to debug
    print("Listing all visible windows:")
    windows = Desktop(backend="win32").windows()
    for w in windows:
        if w.is_visible():
            print(f" - '{w.window_text()}'")

    try:
        print("\nAttempting to connect to 'HiTOPS3' (win32)...")
        # Try finding by exact title first, which seems more reliable
        app = pywinauto.Application(backend="win32").connect(title="HiTOPS3", timeout=10)
        main_dlg = app.top_window()
        print(f"Connected to Window: {main_dlg.window_text()}")
        
        print("\n--- Control Identifiers (Depth 2) ---")
        main_dlg.print_control_identifiers(depth=2)
        
    except Exception as e:
        print(f"Error connecting: {e}")
        try:
             # Fallback to broad regex
             print("\nRetrying with regex '.*Main.*'...")
             app = pywinauto.Application(backend="win32").connect(title_re=".*Main.*", found_index=0, timeout=10)
             main_dlg = app.top_window()
             print(f"Connected to Window: {main_dlg.window_text()}")
             main_dlg.print_control_identifiers(depth=2)
        except Exception as e2:
             print(f"Error connecting fallback: {e2}")

if __name__ == "__main__":
    inspect_uitree()
