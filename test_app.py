import tkinter as tk
from tkinter import messagebox

def test_gui():
    root = tk.Tk()
    root.title("Test Window")
    root.geometry("300x200")
    
    label = tk.Label(root, text="Test Window")
    label.pack(pady=20)
    
    button = tk.Button(root, text="Test Message", command=lambda: messagebox.showinfo("Test", "GUI is working!"))
    button.pack()
    
    root.mainloop()

if __name__ == "__main__":
    print("Starting test application...")
    test_gui() 