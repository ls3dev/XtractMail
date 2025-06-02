import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from pathlib import Path

class ExcelOutlookApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Search & Email")
        self.root.geometry("600x400")
        
        self.df = None
        self.setup_ui()
        
    def setup_ui(self):
        # File selection
        self.file_frame = tk.Frame(self.root)
        self.file_frame.pack(pady=10, padx=10, fill=tk.X)
        
        tk.Button(self.file_frame, text="Select Excel File", command=self.load_excel).pack(side=tk.LEFT)
        self.file_label = tk.Label(self.file_frame, text="No file selected")
        self.file_label.pack(side=tk.LEFT, padx=10)
        
        # Search frame
        self.search_frame = tk.Frame(self.root)
        self.search_frame.pack(pady=10, padx=10, fill=tk.X)
        
        tk.Label(self.search_frame, text="Search Name:").pack(side=tk.LEFT)
        self.search_entry = tk.Entry(self.search_frame)
        self.search_entry.pack(side=tk.LEFT, padx=5)
        tk.Button(self.search_frame, text="Search", command=self.search_data).pack(side=tk.LEFT)
        
        # Results frame
        self.results_frame = tk.Frame(self.root)
        self.results_frame.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)
        
        self.results_text = tk.Text(self.results_frame, height=10)
        self.results_text.pack(fill=tk.BOTH, expand=True)
        
        # Email frame
        self.email_frame = tk.Frame(self.root)
        self.email_frame.pack(pady=10, padx=10, fill=tk.X)
        
        # Email configuration
        tk.Label(self.email_frame, text="SMTP Server:").pack(side=tk.LEFT)
        self.smtp_entry = tk.Entry(self.email_frame)
        self.smtp_entry.pack(side=tk.LEFT, padx=5)
        self.smtp_entry.insert(0, "smtp.gmail.com")
        
        tk.Label(self.email_frame, text="From:").pack(side=tk.LEFT)
        self.from_entry = tk.Entry(self.email_frame)
        self.from_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Label(self.email_frame, text="To:").pack(side=tk.LEFT)
        self.to_entry = tk.Entry(self.email_frame)
        self.to_entry.pack(side=tk.LEFT, padx=5)
        
        tk.Button(self.email_frame, text="Send Email", command=self.send_email).pack(side=tk.LEFT)
        
    def load_excel(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.file_label.config(text=Path(file_path).name)
                messagebox.showinfo("Success", "Excel file loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {str(e)}")
                
    def search_data(self):
        if self.df is None:
            messagebox.showwarning("Warning", "Please load an Excel file first!")
            return
            
        search_term = self.search_entry.get().strip()
        if not search_term:
            messagebox.showwarning("Warning", "Please enter a search term!")
            return
            
        # Search in all columns
        results = []
        for column in self.df.columns:
            matches = self.df[self.df[column].astype(str).str.contains(search_term, case=False, na=False)]
            if not matches.empty:
                results.append(matches)
                
        if results:
            combined_results = pd.concat(results).drop_duplicates()
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, combined_results.to_string())
        else:
            messagebox.showinfo("Info", "No matches found!")
            
    def send_email(self):
        if not self.results_text.get(1.0, tk.END).strip():
            messagebox.showwarning("Warning", "No results to send!")
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
            msg['Subject'] = "Excel Search Results"
            
            body = self.results_text.get(1.0, tk.END)
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

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelOutlookApp(root)
    root.mainloop() 