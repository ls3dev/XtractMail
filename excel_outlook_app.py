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
        self.root = ttk.Window(themename="darkly")
        self.root.title("Excel Viewer & Email")
        self.root.geometry("1000x700")
        self.df = None
        self.date_columns = []
        self.clear_button = None
        self.outlook = None
        self.setup_ui()
        self.initialize_outlook()
        
    def setup_ui(self):
        # Main container with padding
        main_container = ttk.Frame(self.root, padding="20")
        main_container.pack(fill=BOTH, expand=YES)
        
        # Top frame for file operations
        top_frame = ttk.Frame(main_container)
        top_frame.pack(fill=X, pady=(0, 10))
        
        # File selection frame with modern styling
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
        
        # Results frame with improved styling
        results_frame = ttk.LabelFrame(main_container, text="Excel Data", padding="10")
        results_frame.pack(fill=BOTH, expand=YES, pady=(0, 10))
        
        # Create Treeview with scrollbars in a frame
        tree_frame = ttk.Frame(results_frame)
        tree_frame.pack(fill=BOTH, expand=YES, padx=5, pady=5)
        
        # Style configuration for Treeview
        style = ttk.Style()
        style.configure(
            "primary.Treeview",
            rowheight=25,
            background="#2f3136",  # Dark background
            foreground="white",    # Light text
            fieldbackground="#2f3136"  # Dark background for empty space
        )
        style.configure(
            "primary.Treeview.Heading",
            font=("Helvetica", 10, "bold"),
            background="#202225",  # Darker background for headers
            foreground="white"     # Light text for headers
        )
        style.map(
            "primary.Treeview",
            background=[("selected", "#7289da")],  # Discord-like selection color
            foreground=[("selected", "white")]
        )
        
        # Create scrollbars
        v_scrollbar = ttk.Scrollbar(tree_frame, bootstyle="rounded")
        v_scrollbar.pack(side=RIGHT, fill=Y)
        
        h_scrollbar = ttk.Scrollbar(tree_frame, orient=HORIZONTAL, bootstyle="rounded")
        h_scrollbar.pack(side=BOTTOM, fill=X)
        
        # Create Treeview with improved styling
        self.tree = ttk.Treeview(
            tree_frame,
            show="headings",
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            style="primary.Treeview",
            bootstyle="primary"
        )
        self.tree.pack(fill=BOTH, expand=YES)
        
        # Configure scrollbars
        v_scrollbar.config(command=self.tree.yview)
        h_scrollbar.config(command=self.tree.xview)
        
        # Configure column sorting
        self.tree.bind("<Button-1>", self.on_click_column)
        
        # Configure tag for alternating row colors
        self.tree.tag_configure("oddrow", background="#36393f")  # Slightly lighter dark for odd rows
        
        # Email frame with improved layout
        email_frame = ttk.LabelFrame(main_container, text="Email Configuration", padding="10")
        email_frame.pack(fill=X)
        
        # Grid layout for email configuration with better spacing
        email_grid = ttk.Frame(email_frame)
        email_grid.pack(fill=X, padx=10, pady=5)
        
        # Configure grid columns
        email_grid.columnconfigure(1, weight=1)
        
        # Email fields with consistent spacing
        ttk.Label(email_grid, text="To:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky=W)
        self.to_entry = ttk.Entry(email_grid)
        self.to_entry.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky=EW)
        
        ttk.Label(email_grid, text="Subject:").grid(row=1, column=0, padx=(0, 10), pady=5, sticky=W)
        self.subject_entry = ttk.Entry(email_grid)
        self.subject_entry.grid(row=1, column=1, columnspan=2, padx=5, pady=5, sticky=EW)
        self.subject_entry.insert(0, "Excel Data")
        
        ttk.Label(email_grid, text="CC:").grid(row=2, column=0, padx=(0, 10), pady=5, sticky=W)
        self.cc_entry = ttk.Entry(email_grid)
        self.cc_entry.grid(row=2, column=1, columnspan=2, padx=5, pady=5, sticky=EW)
        
        # Message body frame with improved styling
        message_frame = ttk.LabelFrame(email_frame, text="Message", padding="10")
        message_frame.pack(fill=X, padx=10, pady=10)
        
        self.message_text = ttk.Text(message_frame, height=4, width=50)
        self.message_text.pack(fill=X, expand=YES)
        
        # Options frame with better organization
        options_frame = ttk.Frame(email_frame)
        options_frame.pack(fill=X, padx=10, pady=5)
        
        # Checkbuttons with improved styling
        self.attach_excel_var = ttk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Attach Excel File",
            variable=self.attach_excel_var,
            bootstyle="info-round-toggle"
        ).pack(side=LEFT, padx=5)
        
        self.include_table_var = ttk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Include Table in Email",
            variable=self.include_table_var,
            bootstyle="info-round-toggle"
        ).pack(side=LEFT, padx=5)
        
        # Send button with improved styling
        ttk.Button(
            email_frame,
            text="Send via Outlook",
            command=self.send_email,
            bootstyle="success"
        ).pack(pady=10)

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
        # This method is now empty as the outlook initialization logic has been moved to a separate method
        pass

if __name__ == "__main__":
    app = ExcelOutlookApp()
    app.root.mainloop() 