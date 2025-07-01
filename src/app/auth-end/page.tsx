'use client';

import { useEffect, useState, Suspense } from 'react';
import { useSearchParams } from 'next/navigation';

interface TeamsWindow extends Window {
  microsoftTeams?: {
    authentication: {
      notifySuccess: (result: unknown) => void;
      notifyFailure: (reason: string) => void;
    };
  };
}

function AuthEndContent() {
  const searchParams = useSearchParams();
  const [status, setStatus] = useState<'processing' | 'success' | 'error'>('processing');
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const handleAuthResult = async () => {
      try {
        // Get auth code from URL parameters
        const code = searchParams?.get('code');
        const error = searchParams?.get('error');
        const errorDescription = searchParams?.get('error_description');

        if (error) {
          throw new Error(errorDescription || error);
        }

        if (!code) {
          throw new Error('No authorization code received');
        }

        // Exchange the auth code for tokens via our API
        const response = await fetch(`/api/auth/teams?code=${encodeURIComponent(code)}`);

        if (!response.ok) {
          const errorText = await response.text();
          throw new Error(`Token exchange failed: ${response.status} - ${errorText}`);
        }

        // Check if response is a redirect
        if (response.redirected) {
          // Parse the redirect URL to get the user data
          const redirectUrl = new URL(response.url);
          const success = redirectUrl.searchParams.get('success');
          const dataParam = redirectUrl.searchParams.get('data');

          console.log('Redirect URL:', response.url);
          console.log('Success param:', success);
          console.log('Data param exists:', !!dataParam);

          if (success === 'true' && dataParam) {
            try {
              const result = JSON.parse(decodeURIComponent(dataParam));
              setStatus('success');

              // Notify Teams of success
              if (window.parent && window.parent !== window) {
                window.parent.postMessage({
                  type: 'auth-success',
                  result: result
                }, '*');
              }

              // Also try to use Teams SDK if available
              const parentWindow = window.parent as TeamsWindow;
              if (parentWindow && parentWindow.microsoftTeams) {
                parentWindow.microsoftTeams.authentication.notifySuccess(result);
              }
              return;
            } catch (parseError) {
              console.error('Failed to parse data param:', parseError);
              throw new Error(`Failed to parse user data: ${parseError}`);
            }
          } else {
            console.error('Missing success or data params in redirect');
            console.error('Redirect URL params:', Object.fromEntries(redirectUrl.searchParams.entries()));
            throw new Error('Invalid redirect response - missing success or data parameters');
          }
        }

        // Handle JSON response (fallback)
        const result = await response.json();
        setStatus('success');

        // Notify Teams of success
        if (window.parent && window.parent !== window) {
          window.parent.postMessage({
            type: 'auth-success',
            result: result
          }, '*');
        }

        // Also try to use Teams SDK if available
        const parentWindow = window.parent as TeamsWindow;
        if (parentWindow && parentWindow.microsoftTeams) {
          parentWindow.microsoftTeams.authentication.notifySuccess(result);
        }

      } catch (error) {
        console.error('Auth end error:', error);
        setStatus('error');
        setError(error instanceof Error ? error.message : 'Authentication failed');

        // Notify Teams of failure
        if (window.parent && window.parent !== window) {
          window.parent.postMessage({
            type: 'auth-failure',
            error: error instanceof Error ? error.message : 'Authentication failed'
          }, '*');
        }

        // Also try to use Teams SDK if available
        const parentWindow = window.parent as TeamsWindow;
        if (parentWindow && parentWindow.microsoftTeams) {
          parentWindow.microsoftTeams.authentication.notifyFailure(
            error instanceof Error ? error.message : 'Authentication failed'
          );
        }
      }
    };

    handleAuthResult();
  }, [searchParams]);

  if (status === 'processing') {
    return (
      <div className="flex items-center justify-center min-h-screen bg-gray-50">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Completing authentication...</p>
        </div>
      </div>
    );
  }

  if (status === 'error') {
    return (
      <div className="flex items-center justify-center min-h-screen bg-gray-50">
        <div className="text-center">
          <div className="text-red-600 text-2xl mb-4">❌</div>
          <h2 className="text-xl font-semibold text-gray-800 mb-2">Authentication Failed</h2>
          <p className="text-gray-600 mb-4">{error}</p>
          <p className="text-sm text-gray-500">You can close this window and try again.</p>
        </div>
      </div>
    );
  }

  return (
    <div className="flex items-center justify-center min-h-screen bg-gray-50">
      <div className="text-center">
        <div className="text-green-600 text-2xl mb-4">✅</div>
        <h2 className="text-xl font-semibold text-gray-800 mb-2">Authentication Successful</h2>
        <p className="text-gray-600">You can close this window now.</p>
      </div>
    </div>
  );
}

export default function AuthEnd() {
  return (
    <Suspense fallback={
      <div className="flex items-center justify-center min-h-screen bg-gray-50">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-600">Loading...</p>
        </div>
      </div>
    }>
      <AuthEndContent />
    </Suspense>
  );
}