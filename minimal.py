import tkinter as tk
print("Starting minimal test...")

try:
    print("Creating root window...")
    root = tk.Tk()
    print("Root window created")
    
    print("Setting window title...")
    root.title("Minimal Test")
    print("Title set")
    
    print("Setting window size...")
    root.geometry("300x200")
    print("Size set")
    
    print("Creating label...")
    label = tk.Label(root, text="Hello World")
    label.pack()
    print("Label created and packed")
    
    print("Starting main loop...")
    root.mainloop()
    print("Main loop ended")
    
except Exception as e:
    print(f"Error occurred: {str(e)}")
    input("Press Enter to exit...") 