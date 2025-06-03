"""
Microsoft Graph API Configuration

Instructions to get these values:
1. Go to Azure Portal (portal.azure.com)
2. Navigate to Azure Active Directory
3. Go to 'App registrations'
4. Click 'New registration'
5. Name your application
6. Select 'Accounts in any organizational directory (Any Azure AD directory - Multitenant)'
7. For Redirect URI, select 'Public client/native (mobile & desktop)'
8. After registration, copy the Application (client) ID
"""

# Your Azure AD app registration details
CLIENT_ID = "b7e7551e-2627-4f83-b4b8-b09219dce187"  # Your client ID
TENANT_ID = "common"  # Use 'common' to support both organizational and personal accounts

# Authentication settings
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = [
    "User.Read",
    "Mail.Read",
    "MailboxSettings.Read"
]

# API endpoints
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0" 