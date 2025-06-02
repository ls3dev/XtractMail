# Excel Search and Email Application

This application allows users to:
1. Load Microsoft Excel files
2. Search for specific data within the Excel sheets
3. Send the search results via Microsoft Outlook

## Prerequisites

- Python 3.7 or higher
- Microsoft Outlook installed and configured on your system
- The required Python packages (listed in requirements.txt)

## Setup

1. Create a virtual environment (recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Linux/Mac
# OR
.\venv\Scripts\activate  # On Windows
```

2. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python excel_outlook_app.py
```

2. Use the application:
   - Click "Select Excel File" to load your Excel file
   - Enter a search term in the "Search Name" field
   - Click "Search" to find matches
   - Enter recipient email address
   - Click "Send Email" to send the results via Outlook

## Notes

- The application searches across all columns in the Excel sheet
- Search is case-insensitive
- Make sure Outlook is running and properly configured before sending emails
- The application requires proper permissions to access Outlook 