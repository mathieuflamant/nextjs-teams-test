import { NextRequest, NextResponse } from 'next/server';
import jwt from 'jsonwebtoken';
import jwksClient from 'jwks-rsa';

// Type definitions
interface JwtPayload {
  sub: string;
  name?: string;
  email?: string;
  upn?: string;
  iss: string;
  aud: string;
  exp: number;
  iat: number;
}

interface CognitoTokens {
  access_token: string;
  refresh_token: string;
  id_token: string;
  token_type: string;
  expires_in: number;
}

interface TokenExchangeResponse {
  success: boolean;
  message: string;
  user?: {
    sub: string;
    name?: string;
    email?: string;
    upn?: string;
  };
  timestamp: string;
  error?: string;
}

// Microsoft Teams JWKS configuration
const MICROSOFT_JWKS_URI = 'https://login.microsoftonline.com/common/discovery/v2.0/keys';
const MICROSOFT_ISSUER = process.env.MICROSOFT_ISSUER;

// Feature flag for authentication mode
const USE_COGNITO_FEDERATION = process.env.USE_COGNITO_FEDERATION === 'true';

// AWS Cognito configuration
const COGNITO_TOKEN_ENDPOINT = process.env.COGNITO_TOKEN_ENDPOINT;
const COGNITO_CLIENT_ID = process.env.COGNITO_CLIENT_ID;
const COGNITO_CLIENT_SECRET = process.env.COGNITO_CLIENT_SECRET;
const COGNITO_USER_POOL_ID = process.env.COGNITO_USER_POOL_ID;
const COGNITO_REGION = process.env.COGNITO_REGION;
const APP_URL = process.env.NEXT_PUBLIC_APP_URL;
const AZURE_APP_RESOURCE = process.env.NEXT_PUBLIC_AZURE_APP_RESOURCE;
const AZURE_CLIENT_ID = process.env.NEXT_PUBLIC_AZURE_CLIENT_ID;
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET;

// Debug: Log environment variable loading
console.log('Environment variables loaded:', {
  USE_COGNITO_FEDERATION: USE_COGNITO_FEDERATION ? 'ENABLED' : 'DISABLED',
  COGNITO_CLIENT_SECRET: COGNITO_CLIENT_SECRET ? `SET (${COGNITO_CLIENT_SECRET.length} chars)` : 'NOT SET',
  COGNITO_USER_POOL_ID: COGNITO_USER_POOL_ID || 'NOT SET',
  COGNITO_REGION: COGNITO_REGION || 'NOT SET',
  MICROSOFT_ISSUER: MICROSOFT_ISSUER || 'NOT SET',
  AZURE_CLIENT_SECRET: AZURE_CLIENT_SECRET ? `SET (${AZURE_CLIENT_SECRET.length} chars)` : 'NOT SET'
});

// Type assertions (without validation to allow page to load)
const MICROSOFT_ISSUER_VALIDATED = MICROSOFT_ISSUER as string;

console.log('Validated variables debug:', {
  MICROSOFT_ISSUER_VALIDATED: MICROSOFT_ISSUER_VALIDATED || 'NOT SET',
  MICROSOFT_ISSUER_VALIDATED_type: typeof MICROSOFT_ISSUER_VALIDATED
});

const COGNITO_TOKEN_ENDPOINT_VALIDATED = COGNITO_TOKEN_ENDPOINT as string;
const COGNITO_CLIENT_ID_VALIDATED = COGNITO_CLIENT_ID as string;
const COGNITO_CLIENT_SECRET_VALIDATED = COGNITO_CLIENT_SECRET as string;
const COGNITO_USER_POOL_ID_VALIDATED = COGNITO_USER_POOL_ID as string;
const COGNITO_REGION_VALIDATED = COGNITO_REGION as string;

const APP_URL_VALIDATED = APP_URL as string;
const AZURE_CLIENT_ID_VALIDATED = AZURE_CLIENT_ID as string;
const AZURE_CLIENT_SECRET_VALIDATED = AZURE_CLIENT_SECRET as string;

// Initialize JWKS client for Microsoft
const jwksClientInstance = jwksClient({
  jwksUri: MICROSOFT_JWKS_URI,
  cache: true,
  cacheMaxEntries: 5,
  cacheMaxAge: 600000, // 10 minutes
});

// Get signing key for JWT verification
function getKey(header: jwt.JwtHeader, callback: (err: Error | null, key?: string) => void) {
  if (!header.kid) {
    callback(new Error('No key ID in token header'));
    return;
  }
 
  jwksClientInstance.getSigningKey(header.kid, (err: Error | null, key: jwksClient.SigningKey | undefined) => {
    if (err) {
      callback(err);
      return;
    }
    const signingKey = key?.getPublicKey();
    callback(null, signingKey);
  });
}

// Verify Microsoft Teams token
async function verifyTeamsToken(token: string): Promise<JwtPayload> {
  // Validate required environment variables
  if (!MICROSOFT_ISSUER) {
    throw new Error('MICROSOFT_ISSUER environment variable is required');
  }
  if (!AZURE_CLIENT_ID) {
    throw new Error('NEXT_PUBLIC_AZURE_CLIENT_ID environment variable is required');
  }

  console.log('Verifying Teams token:', {
    tokenPreview: token.substring(0, 50) + '...',
    tokenLength: token.length,
    issuer: MICROSOFT_ISSUER_VALIDATED,
    audience: AZURE_CLIENT_ID_VALIDATED
  });

  return new Promise((resolve, reject) => {
    jwt.verify(token, getKey, {
      issuer: MICROSOFT_ISSUER_VALIDATED,
      audience: AZURE_CLIENT_ID_VALIDATED,
      algorithms: ['RS256']
    }, (err, decoded) => {
      if (err) {
        console.error('Token verification failed:', err.message);
        reject(err);
        return;
      }
      if (!decoded || typeof decoded === 'string') {
        console.error('Token verification failed: invalid decoded token');
        reject(new Error('Token verification failed'));
        return;
      }
      console.log('Token verified successfully:', {
        sub: decoded.sub,
        name: decoded.name,
        email: decoded.email,
        upn: decoded.upn,
        iss: decoded.iss,
        aud: decoded.aud,
        exp: decoded.exp,
        iat: decoded.iat
      });
      resolve(decoded as JwtPayload);
    });
  });
}

// Authenticate with Cognito using Teams token as external IdP
async function authenticateWithCognito(teamsToken: string, userEmail: string): Promise<CognitoTokens> {
  // Validate required environment variables
  if (!COGNITO_TOKEN_ENDPOINT) {
    throw new Error('COGNITO_TOKEN_ENDPOINT environment variable is required');
  }
  if (!COGNITO_CLIENT_ID) {
    throw new Error('COGNITO_CLIENT_ID environment variable is required');
  }
  if (!COGNITO_CLIENT_SECRET) {
    throw new Error('COGNITO_CLIENT_SECRET environment variable is required');
  }

    console.log('Authenticating with Cognito using Teams token as external IdP', {
    userEmail: userEmail || 'not provided'
  });

  // Validate required environment variables
  if (!COGNITO_USER_POOL_ID) {
    throw new Error('COGNITO_USER_POOL_ID environment variable is required');
  }
  if (!COGNITO_REGION) {
    throw new Error('COGNITO_REGION environment variable is required');
  }

  console.log('Using real Cognito federation with User Pool:', COGNITO_USER_POOL_ID_VALIDATED);

  // Use Cognito's InitiateAuth API for external provider authentication
  const cognitoEndpoint = `https://cognito-idp.${COGNITO_REGION_VALIDATED}.amazonaws.com/`;

  const authData = {
    AuthFlow: 'ADMIN_USER_PASSWORD_AUTH',
    ClientId: COGNITO_CLIENT_ID_VALIDATED,
    UserPoolId: COGNITO_USER_POOL_ID_VALIDATED,
    AuthParameters: {
      USERNAME: userEmail || 'teams-user',
      PASSWORD: teamsToken, // Using Teams token as password for external auth
      'custom:external_provider': 'AzureAD',
      'custom:external_token': teamsToken
    }
  };

  console.log('Calling Cognito InitiateAuth API:', {
    endpoint: cognitoEndpoint,
    userPoolId: COGNITO_USER_POOL_ID_VALIDATED,
    clientId: COGNITO_CLIENT_ID_VALIDATED,
    authFlow: 'ADMIN_USER_PASSWORD_AUTH'
  });

  const response = await fetch(cognitoEndpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-amz-json-1.1',
      'X-Amz-Target': 'AWSCognitoIdentityProviderService.InitiateAuth'
    },
    body: JSON.stringify(authData)
  });

  if (!response.ok) {
    const errorText = await response.text();
    console.error('Cognito InitiateAuth failed:', {
      status: response.status,
      statusText: response.statusText,
      errorText: errorText
    });
    throw new Error(`Cognito authentication failed: ${response.status} ${errorText}`);
  }

  const cognitoResponse = await response.json();
  console.log('Cognito InitiateAuth response:', {
    hasAuthenticationResult: !!cognitoResponse.AuthenticationResult,
    hasChallenge: !!cognitoResponse.ChallengeName
  });

  // Handle different response types
  if (cognitoResponse.AuthenticationResult) {
    // Successful authentication
    const tokens: CognitoTokens = {
      access_token: cognitoResponse.AuthenticationResult.AccessToken,
      refresh_token: cognitoResponse.AuthenticationResult.RefreshToken,
      id_token: cognitoResponse.AuthenticationResult.IdToken,
      token_type: cognitoResponse.AuthenticationResult.TokenType || 'Bearer',
      expires_in: cognitoResponse.AuthenticationResult.ExpiresIn || 3600
    };
    return tokens;
  } else if (cognitoResponse.ChallengeName) {
    // Handle challenges (like NEW_PASSWORD_REQUIRED, MFA, etc.)
    console.log('Cognito challenge received:', cognitoResponse.ChallengeName);
    throw new Error(`Cognito challenge not implemented: ${cognitoResponse.ChallengeName}`);
  } else {
    // Unexpected response
    console.error('Unexpected Cognito response:', cognitoResponse);
    throw new Error('Unexpected Cognito authentication response');
  }
}

// Set secure session cookie
function setSessionCookie(response: NextResponse, tokens: CognitoTokens) {
  const cookieOptions = {
    httpOnly: true,
    secure: process.env.NODE_ENV === 'production',
    sameSite: 'lax' as const,
    maxAge: 3600, // 1 hour
    path: '/',
  };

  // Set access token cookie
  response.cookies.set('access_token', tokens.access_token, cookieOptions);
  
  // Set refresh token cookie (longer expiry)
  response.cookies.set('refresh_token', tokens.refresh_token, {
    ...cookieOptions,
    maxAge: 30 * 24 * 3600, // 30 days
  });

  // Set ID token cookie
  response.cookies.set('id_token', tokens.id_token, cookieOptions);
}

export async function POST(request: NextRequest): Promise<NextResponse<TokenExchangeResponse>> {
  try {
    const body = await request.json();
    const { token } = body as { token?: string };

    if (!token) {
      return NextResponse.json(
        { success: false, error: 'Teams token is required' } as TokenExchangeResponse,
        { status: 400 }
      );
    }

    // Verify the Teams token using Microsoft JWKS
    console.log('Verifying Teams token...');
    const verifiedToken = await verifyTeamsToken(token);
    console.log('Teams token verified successfully');

    let response: NextResponse<TokenExchangeResponse>;

    if (USE_COGNITO_FEDERATION) {
      // Use Cognito federation
      console.log('Using Cognito federation - exchanging Teams token for Cognito tokens');
      const cognitoTokens = await authenticateWithCognito(token, verifiedToken.email || '');
      console.log('Cognito token exchange completed successfully');

      // Create response with Teams user data
      response = NextResponse.json({
        success: true,
        message: 'Teams-Cognito federation successful',
        user: {
          sub: verifiedToken.sub,
          name: verifiedToken.name,
          email: verifiedToken.email,
          upn: verifiedToken.upn,
        },
        timestamp: new Date().toISOString()
      } as TokenExchangeResponse);

      // Set Cognito session cookies
      setSessionCookie(response, cognitoTokens);
    } else {
      // Use Teams-only authentication
      console.log('Using Teams-only authentication - skipping Cognito token exchange');

      // Create response with Teams user data
      response = NextResponse.json({
        success: true,
        message: 'Teams authentication successful',
        user: {
          sub: verifiedToken.sub,
          name: verifiedToken.name,
          email: verifiedToken.email,
          upn: verifiedToken.upn,
        },
        timestamp: new Date().toISOString()
      } as TokenExchangeResponse);

      // Note: Not setting Cognito session cookies since we're not exchanging tokens
    }

    return response;

  } catch (error) {
    console.error('Token exchange error:', error);
    
    return NextResponse.json(
      { 
        success: false, 
        error: error instanceof Error ? error.message : 'Token exchange failed',
        timestamp: new Date().toISOString()
      } as TokenExchangeResponse,
      { status: 500 }
    );
  }
}

// GET endpoint for authorization code flow (auth-end redirect)
export async function GET(request: NextRequest): Promise<NextResponse> {
  const { searchParams } = new URL(request.url);
  const code = searchParams.get('code');
  const error = searchParams.get('error');
  const errorDescription = searchParams.get('error_description');
  const test = searchParams.get('test');

  // Test endpoint for development
  if (test === 'true') {
    try {
      // Debug: Log all environment variables
      console.log('Environment variables debug:');
      console.log('MICROSOFT_ISSUER:', MICROSOFT_ISSUER ? 'SET' : 'NOT SET');
      console.log('COGNITO_TOKEN_ENDPOINT:', COGNITO_TOKEN_ENDPOINT ? 'SET' : 'NOT SET');
      console.log('COGNITO_CLIENT_ID:', COGNITO_CLIENT_ID ? 'SET' : 'NOT SET');
      console.log('COGNITO_CLIENT_SECRET:', COGNITO_CLIENT_SECRET ? 'SET' : 'NOT SET');
      console.log('APP_URL:', APP_URL ? 'SET' : 'NOT SET');
      console.log('AZURE_APP_RESOURCE:', AZURE_APP_RESOURCE ? 'SET' : 'NOT SET');
      console.log('AZURE_CLIENT_ID:', AZURE_CLIENT_ID ? 'SET' : 'NOT SET');
      console.log('AZURE_CLIENT_SECRET:', AZURE_CLIENT_SECRET ? 'SET' : 'NOT SET');

      // Check if required environment variables are set
      const missingVars = [];
      if (!MICROSOFT_ISSUER) missingVars.push('MICROSOFT_ISSUER');
      if (!COGNITO_TOKEN_ENDPOINT) missingVars.push('COGNITO_TOKEN_ENDPOINT');
      if (!COGNITO_CLIENT_ID) missingVars.push('COGNITO_CLIENT_ID');
      if (!COGNITO_CLIENT_SECRET) missingVars.push('COGNITO_CLIENT_SECRET');
      if (!APP_URL) missingVars.push('NEXT_PUBLIC_APP_URL');
      if (!AZURE_APP_RESOURCE) missingVars.push('NEXT_PUBLIC_AZURE_APP_RESOURCE');
      if (!AZURE_CLIENT_ID) missingVars.push('NEXT_PUBLIC_AZURE_CLIENT_ID');
      if (!AZURE_CLIENT_SECRET) missingVars.push('AZURE_CLIENT_SECRET');

      if (missingVars.length > 0) {
        return NextResponse.json({
          success: false,
          error: `Missing environment variables: ${missingVars.join(', ')}`,
          debug: {
            microsoftIssuer: MICROSOFT_ISSUER ? 'SET' : 'NOT SET',
            cognitoTokenEndpoint: COGNITO_TOKEN_ENDPOINT ? 'SET' : 'NOT SET',
            cognitoClientId: COGNITO_CLIENT_ID ? 'SET' : 'NOT SET',
            cognitoClientSecret: COGNITO_CLIENT_SECRET ? 'SET' : 'NOT SET',
            appUrl: APP_URL ? 'SET' : 'NOT SET',
            azureAppResource: AZURE_APP_RESOURCE ? 'SET' : 'NOT SET',
            azureClientId: AZURE_CLIENT_ID ? 'SET' : 'NOT SET',
            azureClientSecret: AZURE_CLIENT_SECRET ? 'SET' : 'NOT SET'
          },
          user: null,
          timestamp: new Date().toISOString()
        });
      }

      // Create a mock response for testing
      const mockUserData = {
        success: true,
        user: {
          sub: 'test-user-123',
          name: 'Test User',
          email: 'test@example.com',
          upn: 'test@example.com',
        },
        timestamp: new Date().toISOString()
      };

      return NextResponse.json(mockUserData);
    } catch (error) {
      console.error('Test endpoint error:', error);
      return NextResponse.json(
        {
          success: false,
          error: error instanceof Error ? error.message : 'Test failed',
          user: null,
          timestamp: new Date().toISOString()
        },
        { status: 500 }
      );
    }
  }

  if (error) {
    console.error('Authorization error:', error, errorDescription);
    return NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?error=${encodeURIComponent(error)}&error_description=${encodeURIComponent(errorDescription || '')}`);
  }

  if (!code) {
    console.error('No authorization code received');
    return NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?error=no_code&error_description=No authorization code received`);
  }

  try {
    // Debug: Log the token exchange parameters (without exposing the secret)
    console.log('Azure AD token exchange debug:');
    console.log('Client ID:', AZURE_CLIENT_ID_VALIDATED);
    console.log('Client Secret length:', AZURE_CLIENT_SECRET_VALIDATED?.length || 0);
    console.log('Client Secret preview:', AZURE_CLIENT_SECRET_VALIDATED?.substring(0, 4) + '...');
    console.log('Client Secret contains special chars:', /[^a-zA-Z0-9]/.test(AZURE_CLIENT_SECRET_VALIDATED || ''));
    console.log('Code length:', code?.length || 0);
    console.log('Redirect URI:', `${APP_URL_VALIDATED}/auth-end`);

    // Create the request body for Azure AD token exchange
    // Note: URLSearchParams automatically URL-encodes values, but let's ensure proper encoding
    const azureTokenExchangeData = new URLSearchParams();
    azureTokenExchangeData.append('grant_type', 'authorization_code');
    azureTokenExchangeData.append('client_id', AZURE_CLIENT_ID_VALIDATED);
    azureTokenExchangeData.append('client_secret', AZURE_CLIENT_SECRET_VALIDATED);
    azureTokenExchangeData.append('code', code);
    azureTokenExchangeData.append('redirect_uri', `${APP_URL_VALIDATED}/auth-end`);

    // Debug: Log the Azure AD request body (without exposing the secret)
    const azureRequestBody = azureTokenExchangeData.toString();
    console.log('Azure AD request body debug:', {
      bodyLength: azureRequestBody.length,
      bodyPreview: azureRequestBody.substring(0, 200) + '...',
      containsClientSecret: azureRequestBody.includes('client_secret='),
      clientSecretParamLength: azureRequestBody.match(/client_secret=([^&]*)/)?.[1]?.length || 0
    });

    // Exchange authorization code for access token using MICROSOFT_ISSUER as base
    console.log('About to use MICROSOFT_ISSUER_VALIDATED:', {
      value: MICROSOFT_ISSUER_VALIDATED,
      type: typeof MICROSOFT_ISSUER_VALIDATED,
      isUndefined: MICROSOFT_ISSUER_VALIDATED === undefined
    });
    const azureTokenEndpoint = MICROSOFT_ISSUER_VALIDATED.replace('/v2.0', '/oauth2/v2.0/token');
    console.log('Using Azure AD token endpoint:', azureTokenEndpoint);
    
    const tokenResponse = await fetch(azureTokenEndpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: azureRequestBody,
    });

    if (!tokenResponse.ok) {
      const errorText = await tokenResponse.text();
      console.error('Azure AD token exchange failed:', {
        status: tokenResponse.status,
        statusText: tokenResponse.statusText,
        errorText: errorText,
        headers: Object.fromEntries(tokenResponse.headers.entries())
      });
      return NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?error=token_exchange_failed&error_description=${encodeURIComponent(errorText)}`);
    }

    const tokenData = await tokenResponse.json();
    const accessToken = tokenData.access_token;

    // Verify the access token
    const verifiedToken = await verifyTeamsToken(accessToken);

    let response: NextResponse;

    if (USE_COGNITO_FEDERATION) {
      // Use Cognito federation
      console.log('Using Cognito federation in GET endpoint - exchanging Teams token for Cognito tokens');
      const cognitoTokens = await authenticateWithCognito(accessToken, verifiedToken.email || '');
      console.log('Cognito token exchange completed successfully in GET endpoint');

      // Create response with user data
      const userData = {
        success: true,
        user: {
          sub: verifiedToken.sub,
          name: verifiedToken.name,
          email: verifiedToken.email,
          upn: verifiedToken.upn,
        },
        timestamp: new Date().toISOString()
      };

      // Create response with Cognito session cookies
      response = NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?success=true&data=${encodeURIComponent(JSON.stringify(userData))}`);
      setSessionCookie(response, cognitoTokens);
    } else {
      // Use Teams-only authentication
      console.log('Using Teams-only authentication in GET endpoint - skipping Cognito token exchange');

      // Create response with user data
      const userData = {
        success: true,
        user: {
          sub: verifiedToken.sub,
          name: verifiedToken.name,
          email: verifiedToken.email,
          upn: verifiedToken.upn,
        },
        timestamp: new Date().toISOString()
      };

      // Create response without Cognito session cookies
      response = NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?success=true&data=${encodeURIComponent(JSON.stringify(userData))}`);
    }

    return response;

  } catch (error) {
    console.error('Authorization code flow error:', error);
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    return NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?error=processing_failed&error_description=${encodeURIComponent(errorMessage)}`);
  }
}
