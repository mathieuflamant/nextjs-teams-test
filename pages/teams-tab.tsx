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
          const token = await microsoftTeams.authentication.getAuthToken();
          setAuthToken(token);
          
          // Exchange token for Cognito tokens
          await exchangeTokenForCognito(token);
          
        } catch (authError) {
          console.warn("Auth token not available:", authError);
          setError(`Authentication failed: ${authError instanceof Error ? authError.message : 'Unknown error'}. Check Teams app configuration and Azure AD setup.`);
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
        const errorData = await response.json();
        throw new Error(errorData.error || 'Token exchange failed');
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

      // Use Teams authentication flow
      microsoftTeams.authentication.authenticate({
        url: `${process.env.NEXT_PUBLIC_APP_URL}/auth-start`,
        width: 600,
        height: 535,
        successCallback: (result: any) => {
          console.log("Auth success:", result);
          setUserInfo(result.user as UserInfo);
          setTokenExchangeStatus('success');
        },
        failureCallback: (reason: any) => {
          console.error("Auth failed:", reason);
          setTokenExchangeStatus('error');
          setError(`Authentication failed: ${reason}`);
        }
      });

    } catch (error) {
      console.error('Teams authentication error:', error);
      setTokenExchangeStatus('error');
      setError(error instanceof Error ? error.message : 'Authentication failed');
    }
  };

  const testTeamsFunctions = async () => {
    try {
      // Test basic Teams SDK functions
      console.log("Testing Teams SDK functions...");
      
      // Test getting context again
      const currentContext = await microsoftTeams.app.getContext();
      console.log("Current context:", currentContext);
      
    } catch (err) {
      console.error("Teams function test error:", err);
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
          onClick={startTeamsAuthentication}
          className="w-full bg-green-500 hover:bg-green-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
        >
          Start Teams Authentication
        </button>

        <div className="text-xs text-gray-500 text-center">
          This app tests Teams-Cognito token exchange. Requires proper Azure AD and Cognito configuration.
        </div>
      </div>
    </div>
  );
}

