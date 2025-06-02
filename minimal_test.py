import sys
print("Starting minimal test...")
print(f"Python version: {sys.version}")

try:
    print("Importing ttkbootstrap...")
    import ttkbootstrap as ttk
    print("ttkbootstrap imported successfully")
    
    print("Creating window...")
    root = ttk.Window(themename="darkly")
    root.title("Minimal Test")
    root.geometry("300x200")
    
    print("Adding a label...")
    label = ttk.Label(root, text="Test Window")
    label.pack(pady=20)
    
    print("Starting main loop...")
    root.mainloop()
    print("Window closed normally")
    
except Exception as e:
    print(f"Error occurred: {str(e)}")
    input("Press Enter to exit...") 