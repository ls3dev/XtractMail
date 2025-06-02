import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

class ExcelOutlookApp:
    def __init__(self):
        self.root = ttk.Window(themename="superhero")
        self.root.title("Excel Viewer & Email")
        self.root.geometry("1000x700")
        self.df = None
        self.date_columns = []
        self.clear_button = None  # Store reference to clear button
        self.setup_ui()
        
    def setup_ui(self):
        # Main container
        main_container = ttk.Frame(self.root, padding="20")
        main_container.pack(fill=BOTH, expand=YES)
        
        # File selection frame
        file_frame = ttk.LabelFrame(main_container, text="File Selection", padding="10")
        file_frame.pack(fill=X, pady=(0, 10))
        
        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side=LEFT, padx=(0, 10))
        
        ttk.Button(
            file_frame,
            text="Select Excel File",
            command=self.load_excel,
            style="primary.TButton"
        ).pack(side=LEFT, padx=(0, 10))
        
        # Create but don't pack the clear button yet
        self.clear_button = ttk.Button(
            file_frame,
            text="Clear All",
            command=self.clear_all,
            style="danger.TButton"
        )
        
        # Results frame
        results_frame = ttk.LabelFrame(main_container, text="Excel Data", padding="10")
        results_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))
        
        # Create Treeview with scrollbars
        tree_frame = ttk.Frame(results_frame)
        tree_frame.pack(fill=BOTH, expand=YES)
        
        # Create vertical scrollbar
        v_scrollbar = ttk.Scrollbar(tree_frame)
        v_scrollbar.pack(side=RIGHT, fill=Y)
        
        # Create horizontal scrollbar
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=HORIZONTAL)
        h_scrollbar.pack(side=BOTTOM, fill=X)
        
        # Create Treeview
        self.tree = ttk.Treeview(
            tree_frame,
            show="headings",  # Hide the first empty column
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            style="primary.Treeview"
        )
        self.tree.pack(fill=BOTH, expand=YES)
        
        # Configure scrollbars
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)
        
        # Configure column sorting
        self.tree.bind("<Button-1>", self.on_click_column)
        
        # Email frame
        email_frame = ttk.LabelFrame(main_container, text="Email Configuration", padding="10")
        email_frame.pack(fill=X)
        
        # Grid layout for email configuration
        ttk.Label(email_frame, text="SMTP Server:").grid(row=0, column=0, padx=5, pady=5)
        self.smtp_entry = ttk.Entry(email_frame)
        self.smtp_entry.grid(row=0, column=1, padx=5, pady=5, sticky=EW)
        self.smtp_entry.insert(0, "smtp.gmail.com")
        
        ttk.Label(email_frame, text="From:").grid(row=1, column=0, padx=5, pady=5)
        self.from_entry = ttk.Entry(email_frame)
        self.from_entry.grid(row=1, column=1, padx=5, pady=5, sticky=EW)
        
        ttk.Label(email_frame, text="To:").grid(row=2, column=0, padx=5, pady=5)
        self.to_entry = ttk.Entry(email_frame)
        self.to_entry.grid(row=2, column=1, padx=5, pady=5, sticky=EW)
        
        email_frame.columnconfigure(1, weight=1)
        
        ttk.Button(
            email_frame,
            text="Send Email",
            command=self.send_email,
            style="warning.TButton"
        ).grid(row=3, column=0, columnspan=2, pady=10)
        
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
                # Read Excel file with date parsing
                self.df = pd.read_excel(
                    file_path,
                    parse_dates=True,
                    date_parser=lambda x: pd.to_datetime(x, errors='coerce')
                )
                
                self.file_label.config(text=Path(file_path).name)
                
                # Show the clear button after successful file load
                self.clear_button.pack(side=LEFT)
                
                # Detect date columns
                self.detect_date_columns()
                print(f"Detected date columns: {self.date_columns}")  # Debug info
                
                # Clear existing tree items
                for item in self.tree.get_children():
                    self.tree.delete(item)
                
                # Configure columns
                columns = list(self.df.columns)
                self.tree["columns"] = columns
                
                # Set column headings and widths
                for col in columns:
                    self.tree.heading(col, text=str(col), anchor=W)
                    # Calculate column width based on header and data
                    max_width = len(str(col)) * 10
                    for value in self.df[col]:
                        formatted_value = self.format_value(value, col)
                        width = len(str(formatted_value)) * 10
                        if width > max_width:
                            max_width = width
                    self.tree.column(col, width=min(max_width, 300), anchor=W)
                
                # Insert data with formatted dates
                for index, row in self.df.iterrows():
                    formatted_row = [self.format_value(row[col], col) for col in columns]
                    self.tree.insert("", END, values=formatted_row)
                
                messagebox.showinfo("Success", "Excel file loaded successfully!")
            except Exception as e:
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

if __name__ == "__main__":
    app = ExcelOutlookApp()
    app.root.mainloop() 