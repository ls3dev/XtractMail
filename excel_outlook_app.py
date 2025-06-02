import sys
print("Starting application...")
print(f"Python version: {sys.version}")

try:
    print("Importing ttkbootstrap...")
    import ttkbootstrap as ttk
    print("ttkbootstrap imported successfully")
    
    print("Importing other modules...")
    from ttkbootstrap.constants import *
    from tkinter import filedialog, messagebox
    import pandas as pd
    from datetime import datetime
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    from pathlib import Path
    import win32com.client
    import pythoncom
    print("All modules imported successfully")

except Exception as e:
    print(f"Error during imports: {str(e)}")
    input("Press Enter to exit...")
    sys.exit(1)

print("Defining ExcelOutlookApp class...")

class ExcelOutlookApp:
    def __init__(self):
        # Initialize the main window
        self.root = ttk.Window(themename="darkly")
        self.root.title("Excel Viewer & Email")
        self.root.geometry("1000x700")
        
        # Initialize variables
        self.df = None
        self.date_columns = []
        self.clear_button = None
        
        # Setup the basic UI
        self.setup_ui()
        
        # Try to initialize Outlook later
        self.outlook = None
        self.contacts = []
        
    def setup_ui(self):
        # Main container with padding
        main_container = ttk.Frame(self.root, padding="20")
        main_container.pack(fill=BOTH, expand=YES)
        
        # Top frame for file operations
        top_frame = ttk.Frame(main_container)
        top_frame.pack(fill=X, pady=(0, 10))
        
        # File selection frame
        file_frame = ttk.LabelFrame(top_frame, text="File Selection", padding="10")
        file_frame.pack(side=LEFT, fill=X, expand=YES)
        
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=LEFT, padx=(0, 10))
        
        ttk.Button(
            file_frame,
            text="Select Excel File",
            command=self.load_excel,
            bootstyle="info"
        ).pack(side=LEFT, padx=(0, 10))
        
        # Create but don't pack the clear button yet
        self.clear_button = ttk.Button(
            file_frame,
            text="Clear All",
            command=self.clear_all,
            bootstyle="danger"
        )
        
        # Results frame
        results_frame = ttk.LabelFrame(main_container, text="Excel Data", padding="10")
        results_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))
        
        # Create Treeview with scrollbars
        tree_frame = ttk.Frame(results_frame)
        tree_frame.pack(fill=BOTH, expand=YES, padx=5, pady=5)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, bootstyle="rounded")
        v_scrollbar.pack(side=RIGHT, fill=Y)
        
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=HORIZONTAL, bootstyle="rounded")
        h_scrollbar.pack(side=BOTTOM, fill=X)
        
        # Create Treeview
        self.tree = ttk.Treeview(
            tree_frame,
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            bootstyle="primary"
        )
        self.tree.pack(fill=BOTH, expand=YES)
        
        # Configure scrollbars
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)
        
        # Configure tag for alternating row colors
        self.tree.tag_configure("oddrow", background="#36393f")
        
        # Try to initialize Outlook features
        self.try_init_outlook()
        
    def try_init_outlook(self):
        """Try to initialize Outlook integration"""
        try:
            import win32com.client
            import pythoncom
            
            pythoncom.CoInitialize()
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            print("Outlook integration enabled")
            
            # Add Outlook-specific UI elements
            self.setup_email_ui()
            
        except Exception as e:
            print(f"Outlook integration not available: {str(e)}")
            # Continue without Outlook features
            pass
            
    def setup_email_ui(self):
        """Setup email UI elements - only called if Outlook is available"""
        if not self.outlook:
            return
            
        # Add email UI here later
        pass

    def format_value(self, value, column):
        if column in self.date_columns:
            try:
                # Handle pandas Timestamp
                if isinstance(value, pd.Timestamp):
                    return value.strftime('%b-%d')
                # Handle string dates
                if isinstance(value, str):
                    return pd.to_datetime(value).strftime('%b-%d')
                # Handle other date types
                if hasattr(value, 'strftime'):
                    return value.strftime('%b-%d')
                return str(value)
            except:
                return str(value)
        return str(value)

    def detect_date_columns(self):
        self.date_columns = []
        for column in self.df.columns:
            try:
                # Check first non-null value
                sample = self.df[column].dropna().iloc[0] if not self.df[column].empty else None
                if sample is not None:
                    # If it's already a datetime
                    if isinstance(sample, (pd.Timestamp, datetime)):
                        self.date_columns.append(column)
                    # If it's a string, try to parse it
                    elif isinstance(sample, str):
                        try:
                            pd.to_datetime(sample)
                            self.date_columns.append(column)
                        except:
                            pass
            except:
                continue

    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                print(f"\nLoading Excel file: {file_path}")
                # Read Excel file
                self.df = pd.read_excel(
                    file_path,
                    na_filter=True
                )
                
                print(f"Initial columns: {list(self.df.columns)}")
                print(f"Initial shape: {self.df.shape}")
                
                # Filter columns based on non-empty cell count
                valid_columns = self.filter_sparse_columns()
                if valid_columns:
                    # Keep only columns with enough non-empty cells
                    self.df = self.df[valid_columns]
                    print(f"Final shape after filtering: {self.df.shape}")
                else:
                    messagebox.showwarning("Warning", "No columns with sufficient non-empty cells found!")
                    return
                
                self.file_label.config(text=Path(file_path).name)
                
                # Show the clear button after successful file load
                self.clear_button.pack(side=LEFT)
                
                # Detect date columns
                self.detect_date_columns()
                print(f"Detected date columns: {self.date_columns}")
                
                # Clear existing tree items
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                # Configure columns
                columns = list(self.df.columns)
                self.tree["columns"] = columns
                
                # Set column headings and widths with improved styling
                for col in columns:
                    self.tree.heading(col, text=str(col), anchor=W)
                    # Calculate column width based on header and data
                    max_width = len(str(col)) * 10
                    for value in self.df[col]:
                        if pd.notna(value) and value != "":
                            formatted_value = self.format_value(value, col)
                            width = len(str(formatted_value)) * 10
                            if width > max_width:
                                max_width = width
                    self.tree.column(col, width=min(max_width, 300), anchor=W, minwidth=100)
                
                # Insert data with alternating row colors
                for i, row in enumerate(self.df.iterrows()):
                    formatted_row = [self.format_value(row[1][col], col) for col in columns]
                    self.tree.insert("", END, values=formatted_row, tags=('oddrow',) if i % 2 else ())
                
                messagebox.showinfo("Success", f"Excel file loaded successfully!\nKept {len(columns)} columns that have at least 50 non-empty cells.")
            except Exception as e:
                print(f"Error loading file: {str(e)}")
                messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
    
    def on_click_column(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            column = self.tree.identify_column(event.x)
            column_id = self.tree["columns"][int(column[1]) - 1]  # Convert column number to name
            
            # Get current items as a list
            items = [(self.tree.set(item, column_id), item) for item in self.tree.get_children("")]
            
            # Sort items
            items.sort(reverse=getattr(self, "_sort_reverse", False))
            
            # Rearrange items in sorted positions
            for index, (_, item) in enumerate(items):
                self.tree.move(item, "", index)
            
            # Reverse sort next time
            self.tree.heading(column_id, text=f"{column_id} {'↑' if not getattr(self, '_sort_reverse', False) else '↓'}")
            self._sort_reverse = not getattr(self, "_sort_reverse", False)
    
    def send_email(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first!")
            return
            
        smtp_server = self.smtp_entry.get().strip()
        from_email = self.from_entry.get().strip()
        to_email = self.to_entry.get().strip()
        
        if not all([smtp_server, from_email, to_email]):
            messagebox.showwarning("Warning", "Please fill in all email fields!")
            return
            
        try:
            msg = MIMEMultipart()
            msg['From'] = from_email
            msg['To'] = to_email
            msg['Subject'] = "Excel Data"
            
            # Convert DataFrame to string for email
            body = self.df.to_string()
            msg.attach(MIMEText(body, 'plain'))
            
            # Note: For Gmail, you'll need to use an App Password
            password = messagebox.askstring("Password", "Enter your email password:", show='*')
            if not password:
                return
                
            with smtplib.SMTP(smtp_server, 587) as server:
                server.starttls()
                server.login(from_email, password)
                server.send_message(msg)
                
            messagebox.showinfo("Success", "Email sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to send email: {str(e)}")

    def clear_all(self):
        # Clear the DataFrame
        self.df = None
        
        # Reset file label
        self.file_label.config(text="No file selected")
        
        # Clear tree view
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Clear email fields
        self.from_entry.delete(0, END)
        self.to_entry.delete(0, END)
        self.smtp_entry.delete(0, END)
        self.smtp_entry.insert(0, "smtp.gmail.com")  # Reset SMTP server to default
        
        # Reset date columns
        self.date_columns = []
        
        # Hide the clear button
        self.clear_button.pack_forget()
        
        messagebox.showinfo("Success", "All data has been cleared!")

    def filter_sparse_columns(self, min_nonempty=50, sample_size=180):
        """Keep columns that have at least 50 non-empty cells in the first 180 rows."""
        if len(self.df) == 0:
            return []
            
        valid_columns = []
        total_rows = min(len(self.df), sample_size)
        print(f"\nAnalyzing first {total_rows} rows of each column:")
        
        for column in self.df.columns:
            # Count non-empty cells (not NaN and not empty string) in the first sample_size rows
            column_data = self.df[column].head(total_rows)
            nonempty_count = (~column_data.isna() & (column_data != "")).sum()
            print(f"Column '{column}': {nonempty_count} non-empty cells")
            
            if nonempty_count >= min_nonempty:
                valid_columns.append(column)
                print(f"  - Keeping column '{column}' ({nonempty_count} non-empty cells >= {min_nonempty})")
            else:
                print(f"  - Removing column '{column}' (only {nonempty_count} non-empty cells)")
                
        print(f"\nKept {len(valid_columns)} columns out of {len(self.df.columns)}")
        return valid_columns

    def initialize_outlook(self):
        """Initialize Outlook with better error handling"""
        try:
            print("\nInitializing Outlook...")
            pythoncom.CoInitialize()  # Initialize COM for the thread
            self.outlook = win32com.client.Dispatch("Outlook.Application")
            print("Outlook COM object created successfully")
            
            # Test if Outlook is responding
            namespace = self.outlook.GetNamespace("MAPI")
            print("MAPI namespace accessed successfully")
            
            # Load contacts
            self.load_outlook_contacts()
            
        except Exception as e:
            print(f"\nError initializing Outlook: {str(e)}")
            print("Detailed error information:")
            import traceback
            traceback.print_exc()
            
            # Show error but don't crash the application
            messagebox.showwarning(
                "Outlook Warning",
                "Could not initialize Outlook. Email features will be disabled.\n\n"
                f"Error: {str(e)}\n\n"
                "Please ensure:\n"
                "1. Outlook is installed and running\n"
                "2. You have necessary permissions\n"
                "3. You're logged into your Outlook account"
            )
            self.outlook = None
            
    def load_outlook_contacts(self):
        """Load contacts with better error handling"""
        if not self.outlook:
            print("Outlook not initialized, skipping contact loading")
            return
            
        try:
            print("\nLoading Outlook contacts...")
            namespace = self.outlook.GetNamespace("MAPI")
            contacts_folder = namespace.GetDefaultFolder(10)  # 10 is the Contacts folder
            contacts = contacts_folder.Items
            
            self.contacts = []
            contact_count = 0
            error_count = 0
            
            for contact in contacts:
                try:
                    if hasattr(contact, 'Email1Address') and contact.Email1Address:
                        contact_info = {
                            'name': getattr(contact, 'FullName', ''),
                            'email': contact.Email1Address,
                            'company': getattr(contact, 'CompanyName', ''),
                            'department': getattr(contact, 'Department', '')
                        }
                        self.contacts.append(contact_info)
                        contact_count += 1
                        print(f"Loaded contact: {contact_info['name']} ({contact_info['email']})")
                except Exception as contact_error:
                    error_count += 1
                    print(f"Error processing contact: {str(contact_error)}")
                    continue
            
            print(f"\nContact loading complete:")
            print(f"- Successfully loaded {contact_count} contacts")
            if error_count > 0:
                print(f"- Encountered {error_count} errors while loading contacts")
            
            if contact_count > 0:
                messagebox.showinfo("Success", f"Loaded {contact_count} contacts from Outlook")
                # Update the To: field autocomplete
                self.setup_email_autocomplete()
            else:
                messagebox.showwarning(
                    "No Contacts",
                    "No contacts were found in Outlook.\n\n"
                    "Please ensure you have contacts in your Outlook address book."
                )
            
        except Exception as e:
            print(f"\nError loading contacts: {str(e)}")
            print("Detailed error information:")
            import traceback
            traceback.print_exc()
            
            messagebox.showwarning(
                "Contact Loading Error",
                "Could not load Outlook contacts.\n\n"
                f"Error: {str(e)}\n\n"
                "The application will continue without contact features."
            )
            
    def setup_email_autocomplete(self):
        """Setup autocomplete for email fields"""
        if not self.contacts:
            return
            
        # Create a list of email addresses for autocomplete
        email_list = [f"{contact['name']} <{contact['email']}>" for contact in self.contacts]
        
        def autocomplete(event):
            """Handle autocomplete for email fields"""
            widget = event.widget
            current_text = widget.get()
            
            if not current_text:
                return
                
            matches = []
            for email in email_list:
                if current_text.lower() in email.lower():
                    matches.append(email)
            
            if matches:
                # If there's only one match and user pressed Tab, auto-fill it
                if len(matches) == 1 and event.keysym == 'Tab':
                    widget.delete(0, END)
                    widget.insert(0, matches[0])
                    return 'break'  # Prevent default Tab behavior
                
                # Show matches in a popup
                popup = ttk.Toplevel(self.root)
                popup.geometry(f"+{widget.winfo_rootx()}+{widget.winfo_rooty() + widget.winfo_height()}")
                popup.overrideredirect(True)
                
                listbox = ttk.Listbox(popup, bootstyle="dark")
                listbox.pack(fill=BOTH, expand=YES)
                
                for match in matches:
                    listbox.insert(END, match)
                
                def on_select(event):
                    if listbox.curselection():
                        selected = listbox.get(listbox.curselection())
                        widget.delete(0, END)
                        widget.insert(0, selected)
                        popup.destroy()
                
                listbox.bind('<<ListboxSelect>>', on_select)
                listbox.bind('<Escape>', lambda e: popup.destroy())
                
                # Position the popup below the entry widget
                popup.lift()
                
        # Bind autocomplete to email fields
        self.to_entry.bind('<KeyRelease>', autocomplete)
        self.cc_entry.bind('<KeyRelease>', autocomplete)

if __name__ == "__main__":
    try:
        print("Creating application instance...")
        app = ExcelOutlookApp()
        print("Starting main loop...")
        app.root.mainloop()
        print("Application closed normally")
    except Exception as e:
        print(f"Fatal error: {str(e)}")
        input("Press Enter to exit...") 