"""
Simple test to verify Microsoft Graph API connectivity.
First make sure to install required packages:
pip install requests msal

Before running:
1. Copy config.py.template to config.py
2. Update config.py with your Azure AD app registration details
3. Ensure you have the required permissions in Azure Portal
"""

import msal
import requests
import json
import sys
import os
from typing import Optional

# Check for config file
if not os.path.exists('config.py'):
    print("Error: config.py not found!")
    print("Please copy config.py.template to config.py and update the settings.")
    print("See README.md for setup instructions.")
    sys.exit(1)

try:
    from config import CLIENT_ID, TENANT_ID, AUTHORITY, SCOPES
except ImportError as e:
    print(f"Error importing configuration: {e}")
    print("Please ensure config.py contains all required settings.")
    sys.exit(1)

def verify_configuration() -> bool:
    """Verify the configuration settings."""
    if CLIENT_ID == "YOUR_CLIENT_ID":
        print("Error: Please update CLIENT_ID in config.py")
        return False
    
    if not CLIENT_ID or not TENANT_ID:
        print("Error: Missing required configuration.")
        print("Please ensure CLIENT_ID and TENANT_ID are set in config.py")
        return False
    
    return True

def verify_tenant() -> bool:
    """Verify the tenant exists and is accessible."""
    print("\nVerifying tenant access...")
    headers = {
        'Content-Type': 'application/json'
    }
    try:
        response = requests.get(
            f'https://login.microsoftonline.com/{TENANT_ID}/v2.0/.well-known/openid-configuration',
            headers=headers
        )
        if response.status_code == 200:
            tenant_info = response.json()
            print("✓ Tenant verification successful!")
            if TENANT_ID == "organizations":
                print("✓ Multi-tenant configuration detected")
            else:
                print(f"Tenant name: {tenant_info.get('token_endpoint', '').split('/')[3]}")
            return True
        else:
            print("✗ Error verifying tenant:")
            print(f"Status code: {response.status_code}")
            print("Response:", response.text)
            return False
    except Exception as e:
        print("✗ Error verifying tenant:")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        return False

def get_access_token() -> Optional[str]:
    """Get an access token using device code flow."""
    if not verify_configuration():
        return None

    if not verify_tenant():
        return None

    print("\nInitializing authentication...")
    print(f"Authority URL: {AUTHORITY}")
    print(f"Client ID: {CLIENT_ID[:8]}...{CLIENT_ID[-4:]}")
    
    try:
        print("\nCreating MSAL application instance...")
        app = msal.PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY
        )
        print("✓ MSAL application instance created")
        
        # Try the simplest possible scope first
        print("\nTrying authentication with minimal scope (User.Read)...")
        flow = app.initiate_device_flow(scopes=["User.Read"])
        
        if "user_code" not in flow:
            print("\n✗ Failed to create device flow")
            print("Error details:", json.dumps(flow, indent=2))
            print("\nTroubleshooting steps:")
            print("1. Go to Azure Portal > App registrations")
            print("2. Find your application")
            print("3. Under 'Authentication' tab, verify:")
            print("   - Platform is configured as 'Mobile and desktop applications'")
            print("   - Redirect URI includes 'http://localhost'")
            print("   - Allow public client flows is set to 'Yes'")
            print("4. Under 'API permissions' tab, verify:")
            print("   - User.Read permission is added")
            print("   - Admin consent is granted (green check mark)")
            return None
        
        print("\n=== Authentication Required ===")
        print(f"1. Open this URL: {flow['verification_uri']}")
        print(f"2. Enter this code: {flow['user_code']}")
        print("\nWaiting for authentication...")
        
        result = app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            print("✓ Authentication successful!")
            print("\nToken details:")
            print(f"Expires in: {result.get('expires_in', 'unknown')} seconds")
            print(f"Granted scopes: {result.get('scope', 'unknown')}")
            return result["access_token"]
        else:
            print("\n✗ Authentication failed!")
            print("Error type:", result.get("error"))
            print("Error description:", result.get("error_description"))
            print("\nFull error details:", json.dumps(result, indent=2))
            return None
            
    except Exception as e:
        print("\nUnexpected error during authentication:")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        print("\nPlease verify your internet connection and try again.")
        return None

def test_connection():
    print("=== Testing Microsoft Graph API Connection ===")
    print("Using multi-tenant configuration with device code flow...")
    
    # Get token
    token = get_access_token()
    if not token:
        print("\nFailed to get access token. Please check the error messages above.")
        return
    
    print("\nTesting API connection...")
    
    # Test API call
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Accept': 'application/json'
    }
    
    try:
        # First, get organization details to confirm which tenant we're connected to
        print("\n1. Testing organization access...")
        org_response = requests.get(
            'https://graph.microsoft.com/v1.0/organization',
            headers=headers
        )
        
        if org_response.status_code == 200:
            org_data = org_response.json()
            if 'value' in org_data and len(org_data['value']) > 0:
                org = org_data['value'][0]
                print("\n✓ Organization details:")
                print(f"Name: {org.get('displayName')}")
                print(f"Tenant ID: {org.get('id')}")
                print(f"Domain: {org.get('verifiedDomains', [{}])[0].get('name', 'N/A')}")
        else:
            print("✗ Could not fetch organization details")
            print(f"Status code: {org_response.status_code}")
            print("Response:", org_response.text)
        
        # Try to get user profile
        print("\n2. Testing user profile access...")
        response = requests.get(
            'https://graph.microsoft.com/v1.0/me',
            headers=headers
        )
        
        if response.status_code == 200:
            user_data = response.json()
            print("✓ Profile access successful!")
            print(f"Connected as: {user_data.get('displayName')}")
            print(f"Email: {user_data.get('userPrincipalName')}")
            print(f"Account type: {'Guest' if '#EXT#' in user_data.get('userPrincipalName', '') else 'Member'}")
            
            # Print user's roles and permissions if available
            print("\nUser details:")
            print(json.dumps(user_data, indent=2))
        else:
            print("✗ Error accessing profile:")
            print(f"Status code: {response.status_code}")
            print("Response:", response.text)
            return
            
        # Try to get mailbox settings with more detailed error handling
        print("\n3. Testing mailbox settings access...")
        print("Making request to mailbox settings...")
        
        mailbox_response = requests.get(
            'https://graph.microsoft.com/v1.0/me/mailboxSettings',
            headers=headers
        )
        
        print("\nMailbox Settings Response:")
        print(f"Status Code: {mailbox_response.status_code}")
        
        if mailbox_response.status_code != 200:
            print("\n✗ Error accessing mailbox settings:")
            try:
                error_data = mailbox_response.json()
                print("\nDetailed error information:")
                print(json.dumps(error_data, indent=2))
            except:
                print("\nCould not parse error response as JSON. Raw response:")
                print("Raw response text:", mailbox_response.text)
            
            print("\nTroubleshooting steps for multi-tenant setup:")
            print("1. Verify these permissions are granted in Azure Portal:")
            print("   - User.Read")
            print("   - Mail.Read")
            print("   - MailboxSettings.Read")
            print("2. Ensure admin consent is granted in BOTH:")
            print("   - The app's home tenant")
            print("   - The current user's tenant")
            print("3. Check if the authenticated user has required licenses")
            print("4. For guest accounts, verify Exchange Online access is granted")
            return
            
        print("\n✓ Mailbox settings access successful!")
        print(json.dumps(mailbox_response.json(), indent=2))
            
    except Exception as e:
        print("\nUnexpected error during API test:")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        print("Full exception info:", str(sys.exc_info()))

if __name__ == "__main__":
    try:
        test_connection()
    except KeyboardInterrupt:
        print("\nOperation cancelled by user")
    except Exception as e:
        print("\nUnexpected error:")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        print("Full exception info:", str(sys.exc_info()))
    
    print("\nPress Enter to exit...")
    input() 