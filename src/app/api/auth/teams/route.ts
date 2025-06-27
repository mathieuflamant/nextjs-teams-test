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
const MICROSOFT_ISSUER = 'https://login.microsoftonline.com/{tenant-id}/v2.0';

// AWS Cognito configuration
const COGNITO_TOKEN_ENDPOINT = process.env.COGNITO_TOKEN_ENDPOINT || 'https://your-cognito-domain.auth.region.amazoncognito.com/oauth2/token';
const COGNITO_CLIENT_ID = process.env.COGNITO_CLIENT_ID || '';
const COGNITO_CLIENT_SECRET = process.env.COGNITO_CLIENT_SECRET || '';

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
  return new Promise((resolve, reject) => {
    jwt.verify(token, getKey, {
      issuer: MICROSOFT_ISSUER,
      audience: COGNITO_CLIENT_ID,
      algorithms: ['RS256']
    }, (err, decoded) => {
      if (err) {
        reject(err);
        return;
      }
      if (!decoded || typeof decoded === 'string') {
        reject(new Error('Token verification failed'));
        return;
      }
      resolve(decoded as JwtPayload);
    });
  });
}

// Exchange Teams token for Cognito tokens
async function exchangeTokenForCognito(teamsToken: string): Promise<CognitoTokens> {
  const tokenExchangeData = new URLSearchParams({
    grant_type: 'urn:ietf:params:oauth:grant-type:token-exchange',
    subject_token: teamsToken,
    subject_token_type: 'urn:ietf:params:oauth:token-type:id_token',
    client_id: COGNITO_CLIENT_ID,
    client_secret: COGNITO_CLIENT_SECRET,
  });

  const response = await fetch(COGNITO_TOKEN_ENDPOINT, {
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

// GET endpoint for testing/debugging
export async function GET() {
  return NextResponse.json({
    success: true,
    message: 'Teams token exchange endpoint is ready',
    endpoints: {
      microsoft_jwks: MICROSOFT_JWKS_URI,
      cognito_token: COGNITO_TOKEN_ENDPOINT,
    },
    timestamp: new Date().toISOString()
  });
}
