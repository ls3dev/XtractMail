# Microsoft Graph API Test Script

This script helps you test Microsoft Graph API connectivity using your Microsoft 365 account.

## Prerequisites

1. Python 3.6 or higher
2. A Microsoft 365 account
3. Access to Azure Portal (portal.azure.com)

## Setup Instructions

1. Install required packages:
   ```bash
   pip install requests msal
   ```

2. Register your application in Azure Portal:
   1. Go to [Azure Portal](https://portal.azure.com)
   2. Navigate to Azure Active Directory
   3. Go to 'App registrations'
   4. Click 'New registration'
   5. Fill in the details:
      - Name: Choose any name (e.g., "Graph API Test")
      - Supported account types: Select "Accounts in any organizational directory"
      - Redirect URI: Select "Public client/native" and enter "http://localhost"
   6. Click 'Register'
   7. Copy the 'Application (client) ID' - you'll need this later

3. Configure API permissions:
   1. In your app registration, go to 'API permissions'
   2. Click 'Add a permission'
   3. Select 'Microsoft Graph'
   4. Choose 'Delegated permissions'
   5. Add these permissions:
      - User.Read
      - Mail.Read
      - MailboxSettings.Read
   6. Click 'Grant admin consent' (requires admin privileges)

4. Configure the script:
   1. Copy `config.py.template` to `config.py`
   2. Open `config.py` and update:
      - Replace `YOUR_CLIENT_ID` with your Application (client) ID

5. Run the script:
   ```bash
   python simple_test.py
   ```

## Troubleshooting

If you see "account doesn't exist" error:
1. Make sure you're using a Microsoft 365 account
2. Try signing out of all Microsoft accounts in your browser
3. Use an incognito/private browser window

If you get permission errors:
1. Verify all required permissions are added in Azure Portal
2. Ensure admin consent is granted
3. Check if your account has the necessary license (e.g., Exchange Online)

## Support

This script supports:
- Any Microsoft 365 account
- Multi-tenant applications
- Device code flow authentication (no need to handle redirect URIs) 