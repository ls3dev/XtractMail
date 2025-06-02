# Azure AD Configuration
# Replace these values with your Azure AD app registration details
CLIENT_ID = "your-client-id-here"  # Application (client) ID
TENANT_ID = "your-tenant-id-here"  # Directory (tenant) ID

# Authentication settings
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["https://graph.microsoft.com/.default"]

# API endpoints
GRAPH_API_ENDPOINT = "https://graph.microsoft.com/v1.0" 