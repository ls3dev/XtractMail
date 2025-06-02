import sys
import platform
import ctypes
import win32com.client
import pythoncom
import time
import threading
import queue

def print_system_info():
    print("=== System Information ===")
    print(f"Python Version: {sys.version}")
    print(f"Python Architecture: {'64 bit' if sys.maxsize > 2**32 else '32 bit'}")
    print(f"Platform: {platform.platform()}")
    print(f"Process is_64bits: {sys.maxsize > 2**32}")
    print(f"Windows Architecture: {platform.machine()}")
    try:
        is_admin = ctypes.windll.shell32.IsUserAnAdmin()
        print(f"Running as Admin: {bool(is_admin)}")
    except:
        print("Could not determine admin status")
    print("=" * 50 + "\n")

def create_outlook_with_timeout(q):
    try:
        # Initialize COM in this thread
        pythoncom.CoInitialize()
        
        # Try to get existing Outlook instance first
        try:
            print("Attempting to connect to existing Outlook instance...")
            outlook = win32com.client.GetActiveObject("Outlook.Application")
            q.put(("success", outlook))
            return
        except:
            print("No existing Outlook instance found, will create new...")
        
        # Try creating new instance with different methods
        methods = [
            (win32com.client.Dispatch, "Dispatch"),
            (win32com.client.dynamic.Dispatch, "Dynamic Dispatch"),
            (win32com.client.gencache.EnsureDispatch, "EnsureDispatch")
        ]
        
        for method, name in methods:
            try:
                print(f"Trying method: {name}...")
                outlook = method("Outlook.Application")
                # Test if the object is actually working
                version = outlook.Version
                q.put(("success", outlook))
                return
            except Exception as e:
                print(f"Method {name} failed: {str(e)}")
                continue
        
        q.put(("error", "All creation methods failed"))
    except Exception as e:
        q.put(("error", str(e)))
    finally:
        pythoncom.CoUninitialize()

def attempt_outlook_connection():
    print("\nAttempting Outlook Connection with 10 second timeout...")
    
    q = queue.Queue()
    thread = threading.Thread(target=create_outlook_with_timeout, args=(q,))
    thread.daemon = True
    thread.start()
    
    try:
        status, result = q.get(timeout=10)
        if status == "success":
            outlook = result
            print(f"✓ Success! Outlook Version: {outlook.Version}")
            
            print("\nTesting MAPI namespace access...")
            namespace = outlook.GetNamespace("MAPI")
            print("✓ MAPI namespace acquired")
            
            return True
        else:
            print(f"Error: {result}")
            return False
    except queue.Empty:
        print("ERROR: Operation timed out after 10 seconds")
        print("\nThis typically means one of the following:")
        print("1. A security dialog is waiting for user input (but hidden)")
        print("2. Outlook is not responding to COM requests")
        print("3. Another process is blocking Outlook automation")
        return False
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return False

if __name__ == "__main__":
    print_system_info()
    
    print("\nTesting initial COM access...")
    try:
        pythoncom.CoInitialize()
        print("✓ Basic COM initialization successful")
        pythoncom.CoUninitialize()
    except Exception as e:
        print(f"✗ COM initialization failed: {e}")
        sys.exit(1)
    
    success = attempt_outlook_connection()
    if not success:
        print("\nTroubleshooting tips:")
        print("1. Check Task Manager - kill any existing Python processes")
        print("2. Restart Outlook completely")
        print("3. Look for hidden security dialogs")
        print("4. Try running script as Administrator")
        print("5. Check Windows Event Viewer for COM errors")
    
    print("\nPress Enter to exit...")
    input() 