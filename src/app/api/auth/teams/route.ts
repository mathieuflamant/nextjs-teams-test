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

// AWS Cognito configuration
const COGNITO_TOKEN_ENDPOINT = process.env.COGNITO_TOKEN_ENDPOINT;
const COGNITO_CLIENT_ID = process.env.COGNITO_CLIENT_ID;
const COGNITO_CLIENT_SECRET = process.env.COGNITO_CLIENT_SECRET;
const APP_URL = process.env.NEXT_PUBLIC_APP_URL;
const AZURE_APP_RESOURCE = process.env.NEXT_PUBLIC_AZURE_APP_RESOURCE;
const AZURE_CLIENT_ID = process.env.NEXT_PUBLIC_AZURE_CLIENT_ID;

// Type assertions (without validation to allow page to load)
const MICROSOFT_ISSUER_VALIDATED = MICROSOFT_ISSUER as string;
const COGNITO_TOKEN_ENDPOINT_VALIDATED = COGNITO_TOKEN_ENDPOINT as string;
const COGNITO_CLIENT_ID_VALIDATED = COGNITO_CLIENT_ID as string;
const COGNITO_CLIENT_SECRET_VALIDATED = COGNITO_CLIENT_SECRET as string;
const APP_URL_VALIDATED = APP_URL as string;
const AZURE_CLIENT_ID_VALIDATED = AZURE_CLIENT_ID as string;

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

// Exchange Teams token for Cognito tokens
async function exchangeTokenForCognito(teamsToken: string): Promise<CognitoTokens> {
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

  const tokenExchangeData = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:token-exchange',
    subject_token: teamsToken,
    subject_token_type: 'urn:ietf:params:oauth:token-type:id_token',
    client_id: COGNITO_CLIENT_ID_VALIDATED,
    client_secret: COGNITO_CLIENT_SECRET_VALIDATED,
  });

  const response = await fetch(COGNITO_TOKEN_ENDPOINT_VALIDATED, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    body: tokenExchangeData.toString(),
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Cognito token exchange failed: ${response.status} ${errorText}`);
  }

  return await response.json() as CognitoTokens;
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

    // Exchange Teams token for Cognito tokens
    console.log('Exchanging token for Cognito tokens...');
    const cognitoTokens = await exchangeTokenForCognito(token);
    console.log('Token exchange completed successfully');

    // Create response
    const response = NextResponse.json({
      success: true,
      message: 'Token exchange completed successfully',
      user: {
        sub: verifiedToken.sub,
        name: verifiedToken.name,
        email: verifiedToken.email,
        upn: verifiedToken.upn,
      },
      timestamp: new Date().toISOString()
    } as TokenExchangeResponse);

    // Set session cookies
    setSessionCookie(response, cognitoTokens);

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

      // Check if required environment variables are set
      const missingVars = [];
      if (!MICROSOFT_ISSUER) missingVars.push('MICROSOFT_ISSUER');
      if (!COGNITO_TOKEN_ENDPOINT) missingVars.push('COGNITO_TOKEN_ENDPOINT');
      if (!COGNITO_CLIENT_ID) missingVars.push('COGNITO_CLIENT_ID');
      if (!COGNITO_CLIENT_SECRET) missingVars.push('COGNITO_CLIENT_SECRET');
      if (!APP_URL) missingVars.push('NEXT_PUBLIC_APP_URL');
      if (!AZURE_APP_RESOURCE) missingVars.push('NEXT_PUBLIC_AZURE_APP_RESOURCE');
      if (!AZURE_CLIENT_ID) missingVars.push('NEXT_PUBLIC_AZURE_CLIENT_ID');

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
            azureClientId: AZURE_CLIENT_ID ? 'SET' : 'NOT SET'
            allprocessEnvKeys: Object.keys(process.env).sort(),
            allprocessValues: Object.fromEntries(
              Object.entries(process.env).map(([key, value]) => [
                key,
                value || 'undefined'
              ])
            ),
            totalEnvVars: Object.keys(process.env).length
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
    // Exchange authorization code for access token
    const tokenResponse = await fetch('https://login.microsoftonline.com/common/oauth2/v2.0/token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
      },
      body: new URLSearchParams({
        grant_type: 'authorization_code',
        client_id: COGNITO_CLIENT_ID_VALIDATED,
        client_secret: COGNITO_CLIENT_SECRET_VALIDATED,
        code: code,
        redirect_uri: `${APP_URL_VALIDATED}/auth-end`,
      }),
    });

    if (!tokenResponse.ok) {
      const errorText = await tokenResponse.text();
      console.error('Token exchange failed:', errorText);
      return NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?error=token_exchange_failed&error_description=${encodeURIComponent(errorText)}`);
    }

    const tokenData = await tokenResponse.json();
    const accessToken = tokenData.access_token;

    // Verify the access token
    const verifiedToken = await verifyTeamsToken(accessToken);

    // Exchange for Cognito tokens
    const cognitoTokens = await exchangeTokenForCognito(accessToken);

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

    // Create response and set session cookies
    const response = NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?success=true&data=${encodeURIComponent(JSON.stringify(userData))}`);
    setSessionCookie(response, cognitoTokens);

    return response;

  } catch (error) {
    console.error('Authorization code flow error:', error);
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    return NextResponse.redirect(`${APP_URL_VALIDATED}/auth-end?error=processing_failed&error_description=${encodeURIComponent(errorMessage)}`);
  }
}
