import tkinter as tk
from tkinter import messagebox
import threading
import traceback
print("Starting test...")

def init_outlook_thread():
    """Initialize Outlook in a separate thread"""
    print("\nAttempting to initialize Outlook...")
    try:
        print("1. Importing win32com.client...")
        import win32com.client
        print("2. Importing pythoncom...")
        import pythoncom
        
        print("3. Initializing COM...")
        pythoncom.CoInitialize()
        print("COM initialized")
        
        try:
            print("4. Creating Outlook application object...")
            outlook = None  # Define outside try block to check later
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                print("Outlook application object created")
            except Exception as e:
                print(f"Failed to create Outlook object: {str(e)}")
                print("Detailed error:")
                traceback.print_exc()
                raise
            
            if outlook is None:
                raise Exception("Failed to create Outlook object (object is None)")
                
            print("5. Getting MAPI namespace...")
            try:
                namespace = outlook.GetNamespace("MAPI")
                print("Got MAPI namespace")
            except Exception as e:
                print(f"Failed to get MAPI namespace: {str(e)}")
                raise
            
            print("6. Testing inbox access...")
            try:
                inbox = namespace.GetDefaultFolder(6)
                messages = inbox.Items
                msg_count = len(messages)
                print(f"Successfully accessed inbox ({msg_count} messages)")
            except Exception as e:
                print(f"Failed to access inbox: {str(e)}")
                raise
            
            # Get Outlook version
            version = outlook.Version
            print(f"7. Got Outlook version: {version}")
            
            success_msg = (
                f"Outlook Connection Successful!\n\n"
                f"Outlook Version: {version}\n"
                f"Messages in Inbox: {msg_count}\n"
                f"Connection Type: {outlook.Session.ExchangeConnectionMode}"
            )
            
            print("\nSUCCESS:", success_msg)
            root.after(0, lambda: messagebox.showinfo("Success", success_msg))
            root.after(0, lambda: status_label.config(
                text=f"✓ Connected to Outlook {version}",
                fg="green"
            ))
            
        except Exception as inner_e:
            print(f"\nError during Outlook operations: {str(inner_e)}")
            raise
            
        finally:
            print("\nUninitializing COM...")
            pythoncom.CoUninitialize()
            print("COM uninitialized")
            
    except Exception as e:
        error_msg = f"Could not initialize Outlook:\n{str(e)}\n\nPlease ensure:\n1. Outlook is running\n2. You're logged in\n3. You have necessary permissions"
        print(f"\nFINAL ERROR: {error_msg}")
        print("\nFull error trace:")
        traceback.print_exc()
        
        root.after(0, lambda: messagebox.showerror("Error", error_msg))
        root.after(0, lambda: status_label.config(
            text=f"✗ Connection Failed: {str(e)}",
            fg="red"
        ))

def create_window():
    global root, status_label
    print("Creating window...")
    root = tk.Tk()
    root.title("Outlook Test")
    root.geometry("500x400")
    
    def test_outlook():
        # Clear any previous error messages
        print("\n" + "="*50)
        print("Starting new Outlook test...")
        
        # Disable the button while testing
        button.config(state='disabled')
        button.config(text="Testing Connection...")
        status_label.config(text="Testing connection...", fg="black")
        
        # Start Outlook initialization in a separate thread
        thread = threading.Thread(target=init_outlook_thread)
        thread.daemon = True
        thread.start()
        
        # Re-enable button after 3 seconds
        root.after(3000, lambda: button.config(state='normal', text="Test Outlook Connection"))
    
    # Create main frame
    frame = tk.Frame(root, padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    
    # Add instructions
    instructions = tk.Label(
        frame,
        text="This will test the connection to Microsoft Outlook.\n\n"
             "Before testing, please ensure:\n"
             "1. Outlook is installed and running\n"
             "2. You're logged into your Outlook account\n"
             "3. You have permissions to access Outlook programmatically",
        justify=tk.LEFT,
        wraplength=450
    )
    instructions.pack(pady=(0, 20), anchor=tk.W)
    
    # Add the test button
    button = tk.Button(
        frame,
        text="Test Outlook Connection",
        command=test_outlook,
        width=25,
        height=2
    )
    button.pack(pady=(0, 10))
    
    # Add status label with more space for error messages
    status_label = tk.Label(
        frame,
        text="Not connected",
        font=("Arial", 10),
        wraplength=450,
        justify=tk.LEFT
    )
    status_label.pack(pady=10, anchor=tk.W)
    
    print("Starting main loop...")
    root.mainloop()

if __name__ == "__main__":
    try:
        create_window()
        print("Application closed normally")
    except Exception as e:
        print(f"Error: {str(e)}")
        traceback.print_exc()
        input("Press Enter to exit...") 