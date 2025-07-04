# Configuration Guide

## Environment Variables

Create a `.env.local` file in your project root with the following variables:

```bash
# Microsoft Teams Configuration
# Replace {tenant-id} with your actual Azure AD tenant ID
MICROSOFT_ISSUER=https://login.microsoftonline.com/{tenant-id}/v2.0

# AWS Cognito Configuration
# Replace with your actual Cognito domain and region
COGNITO_REGION=us-east-1
COGNITO_EXTERNAL_PROVIDER=MicrosoftEntraID
COGNITO_CLIENT_ID=your-cognito-client-id
COGNITO_CLIENT_SECRET=your-cognito-client-secret

# Teams App Configuration
# This should match your Teams app ID (same as Cognito client ID)
TEAMS_APP_ID=your-teams-app-id
```

## Setup Steps

### 1. Azure AD Configuration
1. Register your application in Azure AD
2. Note your Tenant ID and Application (client) ID
3. Configure redirect URIs for Teams
4. Set up API permissions for Microsoft Graph

### 2. AWS Cognito Configuration
1. Create a Cognito User Pool
2. Create a Cognito App Client
3. Configure the app client for token exchange:
   - Enable "Generate client secret"
   - Set "Allowed OAuth Flows" to include "Authorization code grant"
   - Add your Teams app domain to "Allowed callback URLs"
4. Note your Cognito domain and client credentials

### 3. Teams App Configuration
1. Create a Teams app manifest
2. Set the app ID to match your Cognito client ID
3. Configure the authentication settings
4. Deploy the app to Teams

## Token Exchange Flow

The implemented flow works as follows:

1. **Teams Authentication**: User authenticates in Teams
2. **Token Verification**: Backend verifies Teams ID token using Microsoft JWKS
3. **Token Exchange**: Backend exchanges Teams token for Cognito tokens
4. **Session Setup**: Backend sets secure cookies with Cognito tokens

## Security Considerations

- Store sensitive values in environment variables
- Use HTTPS in production
- Implement proper error handling
- Consider token refresh logic
- Monitor token exchange logs 