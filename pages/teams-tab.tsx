import { useEffect, useState } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";

interface UserInfo {
  sub: string;
  name: string;
  email: string;
  upn: string;
}

export default function TeamsTab() {
  const [isInitialized, setIsInitialized] = useState(false);
  const [authToken, setAuthToken] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [context, setContext] = useState<unknown>(null);
  const [tokenExchangeStatus, setTokenExchangeStatus] = useState<string>('idle');
  const [userInfo, setUserInfo] = useState<UserInfo | null>(null);

  useEffect(() => {
    const initializeTeams = async () => {
      try {
        // Initialize the Teams app
        await microsoftTeams.app.initialize();
        setIsInitialized(true);
        
        // Get the current context
        const context = await microsoftTeams.app.getContext();
        setContext(context);
        
        // Try to get auth token
        try {
          const token = await microsoftTeams.authentication.getAuthToken({
            resources: [process.env.NEXT_PUBLIC_AZURE_APP_RESOURCE!]
          });
          setAuthToken(token);
          // Don't automatically exchange token - let user test manually
          // await exchangeTokenForCognito(token);
        } catch (authError: unknown) {
          const errorMessage = authError instanceof Error ? authError.message : 'Unknown error';

          if (errorMessage.includes('consent_required') || errorMessage.includes('invalid_grant')) {
            console.warn("Consent required, falling back to interactive auth");
            await startTeamsAuthentication();
          } else {
            console.warn("Auth token not available:", errorMessage);
            setError(`Authentication failed: ${errorMessage}.`);
          }
        }
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to initialize Teams');
        console.error("Teams initialization error:", err);
      }
    };

    initializeTeams();
  }, []);

  const exchangeTokenForCognito = async (teamsToken: string) => {
    try {
      setTokenExchangeStatus('exchanging');
      const response = await fetch("/api/auth/teams", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({ token: teamsToken })
      });
      if (!response.ok) {
        const contentType = response.headers.get('content-type');

        if (contentType && contentType.includes('application/json')) {
          const errorData = await response.json();
          throw new Error(errorData.error || 'Token exchange failed');
        } else {
          const errorText = await response.text();
          const errorDetails = `HTTP ${response.status} - Content-Type: ${contentType || 'none'} - Response: ${errorText.substring(0, 100)}...`;
          throw new Error(`API Error: ${errorDetails}`);
        }
      }

      const contentType = response.headers.get('content-type');

      if (!contentType || !contentType.includes('application/json')) {
        const errorText = await response.text();
        const errorDetails = `Expected JSON but got: ${contentType || 'none'} - Response: ${errorText.substring(0, 100)}...`;
        throw new Error(`API Error: ${errorDetails}`);
      }

      const result = await response.json();
      setUserInfo(result.user as UserInfo);
      setTokenExchangeStatus('success');
      console.log('Token exchange successful:', result);
    } catch (error) {
      console.error('Token exchange error:', error);
      setTokenExchangeStatus('error');
      setError(error instanceof Error ? error.message : 'Token exchange failed');
    }
  };

  const startTeamsAuthentication = async () => {
    try {
      setTokenExchangeStatus('exchanging');

      // Construct the auth-start URL with OAuth2 parameters
      const authStartUrl = `${process.env.NEXT_PUBLIC_APP_URL}/auth-start?` +
        `client_id=${encodeURIComponent(process.env.NEXT_PUBLIC_AZURE_CLIENT_ID || '')}` +
        `&redirect_uri=${encodeURIComponent(`${process.env.NEXT_PUBLIC_APP_URL}/auth-end`)}` +
        `&scope=${encodeURIComponent(`openid profile email ${process.env.NEXT_PUBLIC_AZURE_APP_RESOURCE}/access_as_user`)}` +
        `&state=${encodeURIComponent('teams-auth-' + Date.now())}`;

      // Use Teams authentication flow
      const result = await microsoftTeams.authentication.authenticate({
        url: authStartUrl,
        width: 600,
        height: 535,
      });

      // Parse the result which should contain user data
      console.log("Interactive auth success:", result);
      const userData = typeof result === 'string' ? JSON.parse(result) : result;
      setUserInfo(userData.user as UserInfo);
      setTokenExchangeStatus('success');

    } catch (error) {
      console.error('Teams authentication error:', error);
      setTokenExchangeStatus('error');
      setError(error instanceof Error ? error.message : 'Interactive authentication failed');
    }
  };

  const testTeamsFunctions = async () => {
    try {
      // Test basic Teams SDK functions
      console.log("Testing Teams SDK functions...");
      
      // Test getting context again
      const currentContext = await microsoftTeams.app.getContext();
      console.log("Current context:", currentContext);
      
    } catch (error) {
      console.error("Teams function test error:", error);
    }
  };

  const testApiEndpoint = async () => {
    try {
      setTokenExchangeStatus('testing');

      const response = await fetch("/api/auth/teams?test=true");
      const contentType = response.headers.get('content-type');

      if (response.ok) {
        if (contentType && contentType.includes('application/json')) {
          const data = await response.json();

          if (data.success) {
            setTokenExchangeStatus('success');
            setError(null);
            setUserInfo(data.user as UserInfo);
          } else {
            setTokenExchangeStatus('error');
            // Display debug information in the error message
            const debugInfo = data.debug ? `\n\nDebug Info:\n${Object.entries(data.debug).map(([key, value]) => `${key}: ${value}`).join('\n')}` : '';
            setError(`API Test Failed: ${data.error}${debugInfo}`);
          }
        } else {
          const text = await response.text();
          const errorDetails = `API returned non-JSON: ${contentType || 'none'} - Response: ${text.substring(0, 100)}...`;
          setTokenExchangeStatus('error');
          setError(`API Test Failed: ${errorDetails}`);
        }
      } else {
        const text = await response.text();
        const errorDetails = `HTTP ${response.status} - Content-Type: ${contentType || 'none'} - Response: ${text.substring(0, 100)}...`;
        setTokenExchangeStatus('error');
        setError(`API Test Failed: ${errorDetails}`);
      }

    } catch (error) {
      setTokenExchangeStatus('error');
      setError(`API Test Error: ${error instanceof Error ? error.message : 'Unknown error'}`);
    }
  };

  if (error) {
    return (
      <div className="p-6 max-w-md mx-auto bg-white rounded-xl shadow-lg">
        <div className="text-red-600 font-semibold mb-2">Error</div>
        <div className="text-gray-700">{error}</div>
        <div className="text-sm text-gray-500 mt-2">
          This is expected when running outside of Teams environment
        </div>
      </div>
    );
  }

  return (
    <div className="p-6 max-w-2xl mx-auto bg-white rounded-xl shadow-lg">
      <h1 className="text-2xl font-bold text-gray-800 mb-6">
        Teams-Cognito Token Exchange Test
      </h1>
      
      <div className="space-y-4">
        <div className="p-4 bg-blue-50 rounded-lg">
          <h2 className="font-semibold text-blue-800 mb-2">Initialization Status</h2>
          <div className="flex items-center">
            <div className={`w-3 h-3 rounded-full mr-2 ${isInitialized ? 'bg-green-500' : 'bg-yellow-500'}`}></div>
            <span className="text-sm">
              {isInitialized ? 'Teams SDK Initialized' : 'Initializing...'}
            </span>
          </div>
        </div>

        {context !== null && (
          <div className="p-4 bg-green-50 rounded-lg">
            <h2 className="font-semibold text-green-800 mb-2">Teams Context</h2>
            <div className="text-sm text-gray-700">
              <div><strong>Context Available:</strong> Yes</div>
              <div><strong>Context Type:</strong> {typeof context}</div>
            </div>
          </div>
        )}

        <div className="p-4 bg-gray-50 rounded-lg">
          <h2 className="font-semibold text-gray-800 mb-2">Authentication</h2>
          <div className="text-sm text-gray-700">
            <div><strong>Teams Token Available:</strong> {authToken ? 'Yes' : 'No'}</div>
            {authToken && (
              <div className="mt-2">
                <strong>Token Preview:</strong> 
                <div className="bg-gray-100 p-2 rounded text-xs font-mono break-all">
                  {authToken.substring(0, 50)}...
                </div>
              </div>
            )}
          </div>
        </div>

        <div className="p-4 bg-purple-50 rounded-lg">
          <h2 className="font-semibold text-purple-800 mb-2">Token Exchange Status</h2>
          <div className="text-sm text-gray-700">
            <div className="flex items-center mb-2">
              <div className={`w-3 h-3 rounded-full mr-2 ${
                tokenExchangeStatus === 'success' ? 'bg-green-500' :
                tokenExchangeStatus === 'error' ? 'bg-red-500' :
                tokenExchangeStatus === 'exchanging' ? 'bg-yellow-500' :
                'bg-gray-400'
              }`}></div>
              <span className="capitalize">{tokenExchangeStatus}</span>
            </div>
            {userInfo && (
              <div className="mt-2">
                <strong>User Information:</strong>
                <div className="bg-gray-100 p-2 rounded text-xs">
                  <div><strong>Name:</strong> {userInfo.name}</div>
                  <div><strong>Email:</strong> {userInfo.email}</div>
                  <div><strong>UPN:</strong> {userInfo.upn}</div>
                </div>
              </div>
            )}
          </div>
        </div>

        <button
          onClick={testTeamsFunctions}
          className="w-full bg-blue-500 hover:bg-blue-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
        >
          Test Teams Functions
        </button>

        <button
          onClick={testApiEndpoint}
          className="w-full bg-yellow-500 hover:bg-yellow-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
        >
          Test API Endpoint
        </button>

        <button
          onClick={startTeamsAuthentication}
          className="w-full bg-green-500 hover:bg-green-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
        >
          Start Teams Authentication
        </button>

        {authToken && (
          <button
            onClick={() => exchangeTokenForCognito(authToken)}
            className="w-full bg-purple-500 hover:bg-purple-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
          >
            Exchange Token for Cognito
          </button>
        )}

        <div className="text-xs text-gray-500 text-center">
          This app tests Teams-Cognito token exchange. Requires proper Azure AD and Cognito configuration.
        </div>
      </div>
    </div>
  );
}

