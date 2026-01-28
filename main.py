import subprocess
import os
import sys
import login_manager

def launch_hitops():
    """
    Launch Hitops3.exe from the specified path.
    """
    exe_path = r"C:\Program Files (x86)\Hyundai-UNI\HITOPSIII\Hitops3.exe"
    
    if not os.path.exists(exe_path):
        print(f"Error: Executable not found at {exe_path}")
        return

    try:
        # Use subprocess.Popen to launch asynchronously
        working_dir = os.path.dirname(exe_path)
        print(f"Launching {exe_path}...")
        subprocess.Popen(exe_path, cwd=working_dir)
        print("Launch command executed successfully.")
        
        # --- Login Automation ---
        # Note: Ideally, password should be managed securely or input by user.
        # For this version, please replace 'YOUR_PASSWORD_HERE' with your actual password.
        password = "fdjk213!@"  
        login_manager.perform_login(password)
        
    except Exception as e:
        print(f"Failed to launch application: {e}")

if __name__ == "__main__":
    launch_hitops()
