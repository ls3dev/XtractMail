print("=== Starting Simple Outlook Test ===")

print("1. Importing required modules...")
try:
    import win32com.client
    import pythoncom
    import sys
    import time
    print("All modules imported successfully")
except Exception as e:
    print(f"Failed to import modules: {e}")
    input("Press Enter to exit...")
    exit(1)

print("\n2. Checking Python version and architecture...")
print(f"Python Version: {sys.version}")
print(f"Python Architecture: {'64 bit' if sys.maxsize > 2**32 else '32 bit'}")

print("\n3. Initializing COM...")
try:
    pythoncom.CoInitialize()
    print("COM initialized successfully")
except Exception as e:
    print(f"Failed to initialize COM: {e}")
    input("Press Enter to exit...")
    exit(1)

print("\n4. Creating Outlook object...")
outlook = None
errors = []

# First attempt - standard Dispatch
print("Attempting method 1: Standard Dispatch...")
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    print("Success! Outlook object created using standard Dispatch")
except Exception as e:
    errors.append(f"Method 1 failed: {str(e)}")
    print(f"Method 1 failed: {e}")

# Second attempt - GetActiveObject
if outlook is None:
    print("\nAttempting method 2: GetActiveObject...")
    try:
        outlook = win32com.client.GetActiveObject("Outlook.Application")
        print("Success! Connected to running Outlook instance")
    except Exception as e:
        errors.append(f"Method 2 failed: {str(e)}")
        print(f"Method 2 failed: {e}")

# Third attempt - Dynamic Dispatch
if outlook is None:
    print("\nAttempting method 3: Dynamic Dispatch...")
    try:
        outlook = win32com.client.dynamic.Dispatch("Outlook.Application")
        print("Success! Outlook object created using dynamic Dispatch")
    except Exception as e:
        errors.append(f"Method 3 failed: {str(e)}")
        print(f"Method 3 failed: {e}")

if outlook is None:
    print("\nAll connection attempts failed!")
    print("\nDetailed error information:")
    for i, error in enumerate(errors, 1):
        print(f"Attempt {i}: {error}")
    print("\nPossible issues:")
    print("1. Outlook is not installed")
    print("2. Outlook is not running")
    print("3. COM registration issues")
    print("4. Permission/security settings")
    input("Press Enter to exit...")
    exit(1)

print("\n5. Getting MAPI namespace...")
try:
    namespace = outlook.GetNamespace("MAPI")
    print("Got MAPI namespace successfully")
except Exception as e:
    print(f"Failed to get MAPI namespace: {e}")
    input("Press Enter to exit...")
    exit(1)

print("\n6. Getting inbox...")
try:
    inbox = namespace.GetDefaultFolder(6)
    print("Got inbox successfully")
except Exception as e:
    print(f"Failed to get inbox: {e}")
    input("Press Enter to exit...")
    exit(1)

print("\n7. Getting messages...")
try:
    messages = inbox.Items
    count = len(messages)
    print(f"Found {count} messages in inbox")
except Exception as e:
    print(f"Failed to get messages: {e}")
    input("Press Enter to exit...")
    exit(1)

print("\n=== Test Complete ===")
print(f"Successfully connected to Outlook and found {count} messages")
print("\nOutlook Details:")
print(f"Version: {outlook.Version}")
print(f"Connection Mode: {outlook.Session.ExchangeConnectionMode}")

input("\nPress Enter to exit...") 