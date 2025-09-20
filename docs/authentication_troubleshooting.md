# SharePoint Authentication Troubleshooting Guide

This guide helps resolve common authentication issues with SharePoint Excel Manager, particularly when dealing with Conditional Access policies.

## Common Authentication Errors

### AADSTS53003: Access Blocked by Conditional Access
**Error Message:** `AADSTS53003: Access has been blocked by Conditional Access policies. The access policy does not allow token issuance.`

**Cause:** Your organization's Conditional Access policies are blocking the authentication method.

**Solutions:**

1. **Use Device Code Authentication**
   - Click the "Device Auth" button instead of "Test Connection"
   - This method is more compatible with strict Conditional Access policies
   - Follow the instructions in the console to complete authentication

2. **Contact IT Administrator**
   - Your organization may need to configure an app registration specifically for this tool
   - Ask your IT team about Microsoft Graph API access permissions

3. **Use Managed Device**
   - Try running the application from a corporate-managed device
   - Conditional Access policies often allow access from compliant devices

### AADSTS50058: Silent Sign-in Failed
**Error Message:** `AADSTS50058: A silent sign-in request was sent but no user is signed in.`

**Cause:** No cached authentication token available.

**Solution:** Simply try again - this will trigger interactive authentication.

### AADSTS65001: User Consent Required
**Error Message:** `AADSTS65001: The user or administrator has not consented to use the application.`

**Solutions:**

1. **Admin Consent Required**
   - Contact your IT administrator to grant consent for the application
   - The application needs permissions to read SharePoint sites and files

2. **Self-Service Consent**
   - If your organization allows it, you may be able to consent during first login
   - Look for consent prompts during the authentication process

## Authentication Methods

### Method 1: Interactive Authentication (Default)
- Opens a browser window for login
- Best for personal Microsoft 365 accounts
- May be blocked by strict corporate policies

### Method 2: Device Code Authentication
- Displays a code to enter on another device/browser
- Better compatibility with Conditional Access
- Useful for headless or restricted environments
- Click "Device Auth" button to use this method

## Required Permissions

The application requires these Microsoft Graph permissions:

- `Sites.Read.All` - Read SharePoint sites and lists
- `Files.Read.All` - Read files in SharePoint
- `Files.ReadWrite.All` - Upload/modify files (if needed)

## Corporate Environment Setup

### For IT Administrators

If you're setting up this application in a corporate environment:

1. **Create App Registration**
   ```
   - Go to Azure AD > App registrations
   - Create new registration
   - Set redirect URI: http://localhost
   - Grant required Graph API permissions
   - Consider pre-authorizing the application
   ```

2. **Configure Conditional Access**
   ```
   - Create exception for the app registration
   - Allow access from managed devices
   - Configure location-based access if needed
   ```

3. **Update Application Configuration**
   ```python
   # In sharepoint_client.py, update the client_id:
   self.client_id = "your-app-registration-id"
   ```

### For End Users

1. **Check with IT First**
   - Verify if the application is approved for use
   - Ask about any specific authentication requirements

2. **Use Corporate Network**
   - Try connecting from your office network
   - VPN connections may also work

3. **Try Different Authentication Methods**
   - Start with "Test Connection" (interactive)
   - If that fails, try "Device Auth"

## Testing Your Setup

### Quick Test Script

Create a file `test_auth.py`:

```python
import asyncio
from src.sharepoint_excel_manager.sharepoint_client import SharePointClient

async def test_auth():
    client = SharePointClient()
    site_url = input("Enter your SharePoint site URL: ")
    
    print("Testing interactive authentication...")
    success = await client.authenticate(site_url)
    
    if success:
        print("✅ Authentication successful!")
        print(f"Token obtained: {client.access_token[:20]}...")
    else:
        print("❌ Interactive authentication failed")
        print("Trying device code authentication...")
        success = await client.authenticate_device_code(site_url)
        
        if success:
            print("✅ Device code authentication successful!")
        else:
            print("❌ All authentication methods failed")

if __name__ == "__main__":
    asyncio.run(test_auth())
```

Run with: `python test_auth.py`

## Getting Help

If you continue to experience issues:

1. **Check Application Logs**
   - Look for detailed error messages in the console
   - Note the specific error codes (AADSTS*)

2. **Contact IT Support**
   - Provide the specific error codes
   - Mention you're trying to access Microsoft Graph API
   - Ask about Conditional Access policy exceptions

3. **Alternative Solutions**
   - Consider using SharePoint web interface directly
   - Look into Power Platform solutions approved by your organization
   - Use Excel Online with your existing browser authentication

## Security Considerations

- The application uses industry-standard OAuth 2.0 authentication
- No passwords are stored locally
- Tokens are cached temporarily and securely by MSAL
- All communication with SharePoint uses HTTPS
- Consider using managed identities in production environments

## Frequently Asked Questions

**Q: Why does the browser keep opening for authentication?**
A: This is normal for interactive authentication. The browser handles the secure login process.

**Q: Can I skip authentication entirely?**
A: No, authentication is required to access SharePoint resources securely.

**Q: Will this work with personal Microsoft accounts?**
A: Yes, but your organization's SharePoint site must allow external access.

**Q: How long does the authentication last?**
A: Tokens typically last 1 hour, but MSAL automatically refreshes them as needed.

**Q: Is my data secure?**
A: Yes, all authentication follows Microsoft's security standards and your organization's policies.