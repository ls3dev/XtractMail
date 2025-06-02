"""
Simple test to verify Microsoft Graph API connectivity.
First make sure to install required packages:
pip install requests msal
"""

import msal
import requests
import json
import sys

# Your app registration details
CLIENT_ID = "62cdf836-66d1-4f06-9a7d-601335815fbf"
TENANT_ID = "3229079a-de06-4e7b-bdbf-380cbbd0a379"

# Authentication settings
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = ['https://graph.microsoft.com/.default']

def get_access_token():
    print("\nInitializing authentication...")
    print(f"Using tenant ID: {TENANT_ID}")
    print(f"Requesting scopes: {', '.join(SCOPES)}")
    
    try:
        # Initialize the MSAL app as a public client application
        app = msal.PublicClientApplication(
            client_id=CLIENT_ID,
            authority=AUTHORITY
        )
        
        # Clear token cache to ensure fresh authentication
        accounts = app.get_accounts()
        for account in accounts:
            app.remove_account(account)
        
        print("\nStarting new authentication flow...")
        
        # Get token using device code flow
        flow = app.initiate_device_flow(scopes=SCOPES)
        
        if "user_code" not in flow:
            print("\nError: Could not create device flow")
            print("Full error details:", json.dumps(flow, indent=2))
            return None
        
        print("\n=== Authentication Required ===")
        print("1. Open this URL in your browser:", flow["verification_uri"])
        print("2. Enter this code when prompted:", flow["user_code"])
        print("\nWaiting for you to complete the authentication...")
        
        # Wait for user to complete the flow
        result = app.acquire_token_by_device_flow(flow)
        
        if "access_token" in result:
            print("Successfully acquired token!")
            # Print token details for debugging
            print("\nToken details:")
            print(f"Expires in: {result.get('expires_in', 'unknown')} seconds")
            print(f"Granted scopes: {result.get('scope', 'unknown')}")
            print("\nRequested scopes vs Granted scopes:")
            print("Requested:", SCOPES)
            granted_scopes = result.get('scope', '').split(' ')
            print("Granted:", granted_scopes)
            
            # Check if all requested scopes were granted
            missing_scopes = [scope for scope in SCOPES if scope not in granted_scopes]
            if missing_scopes:
                print("\nWARNING: Some requested scopes were not granted:")
                for scope in missing_scopes:
                    print(f"  - {scope}")
            
            return result["access_token"]
        else:
            print("\nError getting token!")
            print("Error type:", result.get("error"))
            print("Error description:", result.get("error_description"))
            print("\nFull error details:", json.dumps(result, indent=2))
            return None
            
    except Exception as e:
        print("\nUnexpected error during authentication:")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        return None

def test_connection():
    print("=== Testing Microsoft Graph API Connection ===")
    print("Using public client authentication with tenant-specific endpoint...")
    
    # Get token
    token = get_access_token()
    if not token:
        print("\nFailed to get access token. Please check the error messages above.")
        return
    
    print("\nTesting API connection...")
    print(f"Token preview (first 50 chars): {token[:50]}...")
    
    # Test API call
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        'Prefer': 'outlook.body-content-type="text"',
        'ConsistencyLevel': 'eventual'
    }
    
    try:
        # Try to get user profile first
        print("\n1. Testing user profile access...")
        response = requests.get(
            'https://graph.microsoft.com/v1.0/me',
            headers=headers
        )
        
        if response.status_code == 200:
            user_data = response.json()
            print("✓ Profile access successful!")
            print(f"Connected as: {user_data.get('displayName')}")
            print(f"Email: {user_data.get('userPrincipalName')}")
            
            # Print user's roles and permissions if available
            print("\nUser details:")
            print(json.dumps(user_data, indent=2))
        else:
            print("✗ Error accessing profile:")
            print(f"Status code: {response.status_code}")
            print("Response:", response.text)
            return
            
        # Try to get mailbox settings with more detailed error handling
        print("\n2. Testing mailbox settings access...")
        print("Trying beta endpoint for more detailed errors...")
        
        # First try beta endpoint
        mailbox_response = requests.get(
            'https://graph.microsoft.com/beta/me/mailboxSettings',
            headers=headers
        )
        
        print("\nMailbox Settings Response (beta):")
        print(f"Status Code: {mailbox_response.status_code}")
        print("Response Headers:", json.dumps(dict(mailbox_response.headers), indent=2))
        
        if mailbox_response.status_code != 200:
            print("\n✗ Error accessing mailbox settings:")
            try:
                error_data = mailbox_response.json()
                print("\nDetailed error information:")
                print(json.dumps(error_data, indent=2))
            except:
                print("\nCould not parse error response as JSON. Raw response:")
                print("Raw response text:", mailbox_response.text)
                
                # Try v1.0 endpoint as fallback
                print("\nTrying v1.0 endpoint as fallback...")
                mailbox_response = requests.get(
                    'https://graph.microsoft.com/v1.0/me/mailboxSettings',
                    headers=headers
                )
                print(f"V1.0 Status Code: {mailbox_response.status_code}")
                try:
                    error_data = mailbox_response.json()
                    print("V1.0 Response:", json.dumps(error_data, indent=2))
                except:
                    print("V1.0 Raw response:", mailbox_response.text)
            
            print("\nTroubleshooting steps:")
            print("1. Verify these permissions are granted in Azure Portal:")
            print("   - Mail.ReadBasic")
            print("   - MailboxSettings.Read")
            print("2. Ensure admin consent is granted for these permissions")
            print("3. Check if the authenticated user has a valid Exchange Online license")
            print("4. Try revoking and re-granting permissions in Azure Portal")
            return
            
        print("\nMailbox settings access successful!")
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