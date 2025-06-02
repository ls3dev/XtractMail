from azure.identity import ClientSecretCredential, DeviceCodeCredential
from msgraph.core import GraphClient
import json
import sys
import traceback
from datetime import datetime
import config

def test_graph_connection():
    print("=== Microsoft Graph API Test ===")
    print(f"Using client ID: {config.CLIENT_ID[:8]}...")  # Show only first 8 chars
    
    try:
        print("\n1. Initializing authentication...")
        # Using DeviceCodeCredential which will display a code to enter
        credential = DeviceCodeCredential(
            client_id=config.CLIENT_ID,
            tenant_id=config.TENANT_ID,
            callback=lambda code: print(f"\nTo sign in, use a web browser to open {code.verification_uri} and enter the code {code.user_code}")
        )
        
        # Create the Graph client
        print("\n2. Creating Graph client...")
        graph_client = GraphClient(credential=credential, scopes=config.SCOPE)
        
        # Test connection by getting user profile
        print("\n3. Testing connection (getting user profile)...")
        response = graph_client.get('/me')
        user_data = response.json()
        
        print("\n=== Connection Successful! ===")
        print(f"Connected as: {user_data.get('displayName', 'Unknown')}")
        print(f"Email: {user_data.get('userPrincipalName', 'Unknown')}")
        
        # Test email access
        print("\n4. Testing email access...")
        response = graph_client.get('/me/messages?$top=1&$select=subject,receivedDateTime')
        email_data = response.json()
        
        if 'value' in email_data and len(email_data['value']) > 0:
            latest_email = email_data['value'][0]
            print("\nLatest email:")
            print(f"Subject: {latest_email.get('subject', 'No subject')}")
            print(f"Received: {latest_email.get('receivedDateTime', 'Unknown')}")
        else:
            print("No emails found or no access to emails")
        
        return True
        
    except Exception as e:
        print("\n=== Error ===")
        print(f"Error type: {type(e).__name__}")
        print(f"Error message: {str(e)}")
        print("\nFull traceback:")
        traceback.print_exc()
        return False

if __name__ == "__main__":
    print("Starting Microsoft Graph API test...")
    print("Note: You will be given a code to enter in your web browser.")
    print("This is a more reliable authentication method.\n")
    
    success = test_graph_connection()
    
    if not success:
        print("\nTroubleshooting tips:")
        print("1. Verify your Azure AD application settings:")
        print(f"   - Check if Client ID is correct: {config.CLIENT_ID[:8]}...")
        print(f"   - Check if Tenant ID is correct: {config.TENANT_ID[:8]}...")
        print("2. Ensure the application has the required permissions:")
        print("   - User.Read")
        print("   - Mail.Read")
        print("3. Make sure you've granted admin consent for these permissions")
        print("4. Check your internet connection")
    
    input("\nPress Enter to exit...") 