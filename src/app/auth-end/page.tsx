'use client';

import { useEffect, useState } from 'react';
import { useSearchParams } from 'next/navigation';

export default function AuthEnd() {
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
        const response = await fetch('/api/auth/teams', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
          },
          body: JSON.stringify({
            code,
            grant_type: 'authorization_code'
          })
        });

        if (!response.ok) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Token exchange failed');
        }

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
        if (window.parent && (window.parent as any).microsoftTeams) {
          (window.parent as any).microsoftTeams.authentication.notifySuccess(result);
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
        if (window.parent && (window.parent as any).microsoftTeams) {
          (window.parent as any).microsoftTeams.authentication.notifyFailure(
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