import ttkbootstrap as ttk
print("Starting ttkbootstrap test...")

try:
    print("Creating window...")
    root = ttk.Window(themename="darkly")
    print("Window created")
    
    print("Setting window properties...")
    root.title("Bootstrap Test")
    root.geometry("300x200")
    print("Properties set")
    
    print("Creating label...")
    label = ttk.Label(root, text="Hello Bootstrap!")
    label.pack(pady=20)
    print("Label created")
    
    print("Creating button...")
    button = ttk.Button(root, text="Test Button", bootstyle="info")
    button.pack()
    print("Button created")
    
    print("Starting main loop...")
    root.mainloop()
    print("Main loop ended")
    
except Exception as e:
    print(f"Error occurred: {str(e)}")
    input("Press Enter to exit...") 