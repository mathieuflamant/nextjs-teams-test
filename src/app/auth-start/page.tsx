'use client';

import { useEffect, Suspense } from 'react';
import { useSearchParams } from 'next/navigation';

function AuthStartContent() {
  const searchParams = useSearchParams();

  useEffect(() => {
    const initiateAuth = async () => {
      try {
        // Get parameters from the URL
        const clientId = searchParams?.get('client_id');
        const redirectUri = searchParams?.get('redirect_uri');
        const scope = searchParams?.get('scope');
        const state = searchParams?.get('state');

        if (!clientId || !redirectUri) {
          console.error('Missing required parameters');
          return;
        }

        // Construct the Azure AD authorization URL
        const authUrl = `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?` +
          `client_id=${encodeURIComponent(clientId)}` +
          `&response_type=code` +
          `&redirect_uri=${encodeURIComponent(redirectUri)}` +
          `&scope=${encodeURIComponent(scope || 'openid profile email')}` +
          `&state=${encodeURIComponent(state || '')}` +
          `&prompt=consent`;

        // Redirect to Azure AD for authentication
        window.location.href = authUrl;

      } catch (error) {
        console.error('Auth start error:', error);
        // Notify Teams of failure
        if (window.parent && window.parent !== window) {
          window.parent.postMessage({
            type: 'auth-failure',
            error: error instanceof Error ? error.message : 'Authentication failed'
          }, '*');
        }
      }
    };

    initiateAuth();
  }, [searchParams]);

  return (
    <div className="flex items-center justify-center min-h-screen bg-gray-50">
      <div className="text-center">
        <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
        <p className="text-gray-600">Initiating authentication...</p>
      </div>
    </div>
  );
}

export default function AuthStart() {
  return (
    <Suspense fallback={
      <div className="flex items-center justify-center min-h-screen bg-gray-50">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Loading...</p>
        </div>
      </div>
    }>
      <AuthStartContent />
    </Suspense>
  );
}