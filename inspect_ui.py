from pywinauto import Desktop
import time
import sys

def inspect_app():
    print("Connecting to Hitops3 application...")
    
    # Try to find the window
    # Search for "HITOPS" in title (based on previous findings)
    try:
        app_ref = Desktop(backend="uia").window(title_re=".*HITOPS.*")
        
        if not app_ref.exists():
            print("Could not find HITOPS window using UIA backend. Trying win32 backend...")
            app_ref = Desktop(backend="win32").window(title_re=".*HITOPS.*")
            
        if not app_ref.exists():
            print("Error: Could not find HITOPS window. Please ensure the application is running.")
            return

        print(f"Found Window: {app_ref.window_text()}")
        print("-" * 50)
        print("Dumping Control Identifiers (this might take a moment)...")
        print("-" * 50)
        
        # Dump identifiers to file to avoid console spam and encoding issues
        try:
            output_file = "ui_dump.txt"
            # Get the control structure text
            # We redirect stdout or just use the method's print capability
            # pywinauto's print_control_identifiers prints to stdout by default
            # We will capture it
            
            from io import StringIO
            old_stdout = sys.stdout
            result = StringIO()
            sys.stdout = result
            
            app_ref.print_control_identifiers()
            
            sys.stdout = old_stdout
            dump_text = result.getvalue()
            
            with open(output_file, "w", encoding="utf-8") as f:
                f.write(dump_text)
                
            print(f"UI Dump saved to: {output_file}")
            print("Please check this file for 'Maintenance' or 'Repair' text.")
            
            # Quick check for keywords
            print("-" * 50)
            print("Analysis:")
            keywords = ["Maintenance", "Repair", "M&R", "Code", "Planning", "Billing"]
            found = []
            for kw in keywords:
                if kw in dump_text:
                    found.append(kw)
            
            if found:
                print(f"SUCCESS: Found text keywords directly in UI: {found}")
                print("We can likely use text recognition!")
            else:
                print("WARNING: Could not find specific text keywords. The UI might be custom-drawn (images).")
                print("We might need to fallback to Image Recognition.")
                
        except Exception as e:
            sys.stdout = sys.__stdout__ # Restore standard out
            print(f"Error during inspection: {e}")

    except Exception as e:
        print(f"Critical Error: {e}")

if __name__ == "__main__":
    inspect_app()
