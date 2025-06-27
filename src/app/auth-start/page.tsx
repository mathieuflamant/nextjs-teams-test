'use client';

import { useEffect, Suspense } from 'react';
import { useSearchParams } from 'next/navigation';

function AuthStartContent() {
  const searchParams = useSearchParams();

  useEffect(() => {
    const initiateAuth = async () => {
      try {
        // Get the auth URL from query parameters (Teams will provide this)
        const authUrl = searchParams?.get('authUrl');

        if (!authUrl) {
          console.error('No auth URL provided');
          return;
        }

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