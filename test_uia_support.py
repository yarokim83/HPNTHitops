"""
UI Automation Support Test for HI-TOPS III
Tests whether the application supports pywinauto/UIA for advanced automation.
Run this while HI-TOPS is already open and logged in.
"""

def test_uia_support():
    """
    Tests if HI-TOPS III supports UI Automation.
    Returns True if supported, False otherwise.
    """
    try:
        from pywinauto import Desktop
        print("✓ pywinauto installed successfully")
    except ImportError:
        print("✗ pywinauto not installed!")
        print("  Install with: pip install pywinauto")
        return False
    
    print("\n=== Searching for HI-TOPS window ===")
    
    # Try UIA backend (modern)
    try:
        desktop = Desktop(backend="uia")
        windows = desktop.windows()
        
        hitops_found = False
        for window in windows:
            title = window.window_text()
            if "HiTOPS" in title or "HI-TOPS" in title:
                print(f"\n✓ Found HI-TOPS window (UIA): '{title}'")
                hitops_found = True
                
                # Try to list controls
                print("\n=== Trying to enumerate UI controls ===")
                try:
                    window.print_control_identifiers(depth=2)
                    print("\n✓ SUCCESS: HI-TOPS supports UI Automation!")
                    print("  → Phase 2 (UIA conversion) is VIABLE")
                    return True
                except Exception as e:
                    print(f"\n✗ Failed to enumerate controls: {e}")
                    print("  HI-TOPS window found but UIA access limited")
                break
        
        if not hitops_found:
            print("\n✗ HI-TOPS window not found with UIA backend")
            print("  Make sure HI-TOPS is running and logged in")
            
    except Exception as e:
        print(f"\n✗ UIA backend error: {e}")
    
    # Try Win32 backend (legacy fallback)
    print("\n=== Trying Win32 backend (legacy) ===")
    try:
        desktop = Desktop(backend="win32")
        windows = desktop.windows()
        
        for window in windows:
            title = window.window_text()
            if "HiTOPS" in title or "HI-TOPS" in title:
                print(f"\n✓ Found HI-TOPS window (Win32): '{title}'")
                
                try:
                    window.print_control_identifiers(depth=1)
                    print("\n⚠ HI-TOPS accessible via Win32 backend (legacy)")
                    print("  → UIA not supported, but Win32 automation possible")
                    print("  → Phase 2 can proceed with Win32 backend")
                    return True
                except Exception as e:
                    print(f"\n✗ Failed to enumerate controls: {e}")
                break
                
    except Exception as e:
        print(f"\n✗ Win32 backend error: {e}")
    
    print("\n=== RESULT ===")
    print("✗ HI-TOPS does not support pywinauto automation")
    print("  → Phase 2 NOT viable, continue with Phase 1 optimization")
    return False

def interactive_test():
    """
    Interactive mode: lets user explore HI-TOPS UI elements.
    """
    try:
        from pywinauto import Application
        
        print("\n=== Interactive UI Exploration ===")
        print("Connecting to HI-TOPS...")
        
        # Try to connect to running process
        app = Application(backend="uia").connect(title_re=".*HiTOPS.*", timeout=5)
        main_window = app.window(title_re=".*HiTOPS.*")
        
        print(f"Connected to: {main_window.window_text()}")
        print("\nFull control tree:")
        main_window.print_control_identifiers(depth=5)
        
        return True
        
    except Exception as e:
        print(f"Interactive test failed: {e}")
        print("Falling back to basic test...")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("HI-TOPS III - UI Automation Support Test")
    print("=" * 60)
    print("\nPrerequisites:")
    print("  1. HI-TOPS must be running and logged in")
    print("  2. Main window should be visible")
    print("\nStarting tests...\n")
    
    # Run basic test
    supported = test_uia_support()
    
    # If supported, offer interactive exploration
    if supported:
        print("\n" + "=" * 60)
        response = input("\nRun interactive exploration? (y/n): ")
        if response.lower() == 'y':
            interactive_test()
    
    print("\n" + "=" * 60)
    print("Test complete. Press Enter to exit...")
    input()
