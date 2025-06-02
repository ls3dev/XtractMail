import win32com.client
import pythoncom

def test_outlook():
    print("Starting Outlook test...")
    
    try:
        print("Initializing COM...")
        pythoncom.CoInitialize()
        print("COM initialized")
        
        print("Creating Outlook object...")
        outlook = win32com.client.Dispatch("Outlook.Application")
        print("Outlook object created")
        
        print("Getting MAPI namespace...")
        namespace = outlook.GetNamespace("MAPI")
        print("Got MAPI namespace")
        
        print("Getting contacts folder...")
        contacts_folder = namespace.GetDefaultFolder(10)
        print("Got contacts folder")
        
        print("Getting contacts...")
        contacts = contacts_folder.Items
        print(f"Found {len(contacts)} contacts")
        
        print("\nTest completed successfully!")
        return True
        
    except Exception as e:
        print(f"\nError occurred: {str(e)}")
        return False

if __name__ == "__main__":
    print("=== Outlook Connection Test ===")
    success = test_outlook()
    print("\nResult:", "Success" if success else "Failed")
    input("Press Enter to exit...") 