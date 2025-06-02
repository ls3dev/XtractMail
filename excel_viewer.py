print("=== Starting Application ===")
print("1. Importing basic modules...")
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
from pathlib import Path
print("Basic modules imported successfully")

print("\n2. Checking for Outlook modules...")
OUTLOOK_AVAILABLE = False
try:
    import win32com.client
    import pythoncom
    OUTLOOK_AVAILABLE = True
    print("Outlook modules imported successfully")
except ImportError as e:
    print(f"Outlook modules not available: {str(e)}")
except Exception as e:
    print(f"Unexpected error importing Outlook modules: {str(e)}")

print("\n3. Defining ExcelViewer class...")

class ExcelViewer:
    def __init__(self):
        print("\n4. Initializing ExcelViewer...")
        try:
            print("4.1 Creating main window...")
            self.root = ttk.Window(themename="darkly")
            print("4.2 Setting window properties...")
            self.root.title("Excel Viewer")
            self.root.geometry("1000x700")
            
            print("4.3 Initializing variables...")
            self.df = None
            self.date_columns = []
            self.clear_button = None
            self.outlook = None
            
            print("4.4 Setting up UI...")
            self.setup_ui()
            print("4.5 UI setup complete")
            
            if OUTLOOK_AVAILABLE:
                print("4.6 Attempting Outlook initialization...")
                self.try_init_outlook()
            else:
                print("4.6 Skipping Outlook initialization (not available)")
                
            print("4.7 Initialization complete")
            
        except Exception as e:
            print(f"ERROR in initialization: {str(e)}")
            raise
        
    def setup_ui(self):
        try:
            print("5.1 Creating main container...")
            main_container = ttk.Frame(self.root, padding="20")
            main_container.pack(fill=BOTH, expand=YES)
            
            print("5.2 Setting up file frame...")
            file_frame = ttk.LabelFrame(main_container, text="File Selection", padding="10")
            file_frame.pack(fill=X, pady=(0, 10))
            
            self.file_label = ttk.Label(file_frame, text="No file selected")
            self.file_label.pack(side=LEFT, padx=(0, 10))
            
            ttk.Button(
                file_frame,
                text="Select Excel File",
                command=self.load_excel,
                bootstyle="info"
            ).pack(side=LEFT, padx=(0, 10))
            
            self.clear_button = ttk.Button(
                file_frame,
                text="Clear",
                command=self.clear_all,
                bootstyle="danger"
            )
            print("5.3 File frame setup complete")
            
            print("5.4 Setting up data frame...")
            data_frame = ttk.LabelFrame(main_container, text="Excel Data", padding="10")
            data_frame.pack(fill=BOTH, expand=YES)
            
            tree_frame = ttk.Frame(data_frame)
            tree_frame.pack(fill=BOTH, expand=YES, padx=5, pady=5)
            
            v_scrollbar = ttk.Scrollbar(tree_frame, bootstyle="rounded")
            v_scrollbar.pack(side=RIGHT, fill=Y)
            
            h_scrollbar = ttk.Scrollbar(tree_frame, orient=HORIZONTAL, bootstyle="rounded")
            h_scrollbar.pack(side=BOTTOM, fill=X)
            
            self.tree = ttk.Treeview(
                tree_frame,
                show="headings",
                yscrollcommand=v_scrollbar.set,
                xscrollcommand=h_scrollbar.set,
                bootstyle="primary"
            )
            self.tree.pack(fill=BOTH, expand=YES)
            
            v_scrollbar.config(command=self.tree.yview)
            h_scrollbar.config(command=self.tree.xview)
            
            self.tree.tag_configure("oddrow", background="#36393f")
            print("5.5 Data frame setup complete")
            
        except Exception as e:
            print(f"ERROR in UI setup: {str(e)}")
            raise
        
    def try_init_outlook(self):
        try:
            print("6.1 Initializing COM...")
            pythoncom.CoInitialize()
            
            print("6.2 Creating Outlook object...")
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            
            print("6.3 Testing MAPI namespace...")
            namespace = self.outlook.GetNamespace("MAPI")
            
            print("6.4 Adding email button...")
            self.add_email_button()
            
            print("6.5 Outlook initialization complete")
            
        except Exception as e:
            print(f"ERROR in Outlook initialization: {str(e)}")
            self.outlook = None

    def add_email_button(self):
        """Add a simple email button to test Outlook functionality"""
        if not self.outlook:
            return
            
        print("Adding email button...")
        email_frame = ttk.Frame(self.root)
        email_frame.pack(fill=X, padx=20, pady=(0, 20))
        
        ttk.Button(
            email_frame,
            text="Test Outlook Connection",
            command=self.test_outlook,
            bootstyle="info"
        ).pack(side=LEFT)
        
    def test_outlook(self):
        """Test the Outlook connection"""
        try:
            namespace = self.outlook.GetNamespace("MAPI")
            inbox = namespace.GetDefaultFolder(6)  # 6 is the Inbox folder
            messages = inbox.Items
            count = len(messages)
            
            messagebox.showinfo(
                "Outlook Test",
                f"Successfully connected to Outlook!\nFound {count} messages in inbox."
            )
            
        except Exception as e:
            print(f"Outlook test failed: {str(e)}")
            messagebox.showerror(
                "Outlook Error",
                f"Could not access Outlook: {str(e)}\n\n"
                "Please ensure Outlook is running and you have necessary permissions."
            )

    def load_excel(self):
        print("Loading Excel file...")
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if file_path:
            try:
                print(f"Reading file: {file_path}")
                self.df = pd.read_excel(file_path)
                print(f"File loaded, shape: {self.df.shape}")
                
                # Update UI
                self.file_label.config(text=Path(file_path).name)
                self.clear_button.pack(side=LEFT)
                
                # Clear existing tree items
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                # Configure columns
                columns = list(self.df.columns)
                self.tree["columns"] = columns
                
                # Set column headings
                for col in columns:
                    self.tree.heading(col, text=str(col))
                    max_width = len(str(col)) * 10
                    self.tree.column(col, width=min(max_width, 300), minwidth=100)
                
                # Insert data
                for i, row in enumerate(self.df.iterrows()):
                    values = [str(val) for val in row[1]]
                    self.tree.insert("", END, values=values, tags=('oddrow',) if i % 2 else ())
                
                print("Data loaded into tree view")
                messagebox.showinfo("Success", "Excel file loaded successfully!")
                
            except Exception as e:
                print(f"Error loading file: {str(e)}")
                messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
    
    def clear_all(self):
        print("Clearing all data...")
        self.df = None
        self.file_label.config(text="No file selected")
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.clear_button.pack_forget()
        print("All data cleared")

print("\n7. Creating application instance...")
try:
    app = ExcelViewer()
    print("\n8. Starting main loop...")
    app.root.mainloop()
    print("\n9. Application closed normally")
except Exception as e:
    print(f"\nFATAL ERROR: {str(e)}")
    import traceback
    traceback.print_exc()
    input("\nPress Enter to exit...") 