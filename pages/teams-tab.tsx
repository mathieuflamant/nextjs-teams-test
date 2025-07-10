import { useEffect, useState } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";
import * as AdaptiveCards from "adaptivecards";

interface UserInfo {
  sub: string;
  name: string;
  email: string;
  upn: string;
}

interface TeamsContext {
  user?: {
    id?: string;
    displayName?: string;
    userPrincipalName?: string;
    email?: string;
  };
  team?: {
    id?: string;
    displayName?: string;
    internalId?: string;
  };
  channel?: {
    id?: string;
    displayName?: string;
  };
  app?: {
    id?: string;
    sessionId?: string;
  };
  locale?: string;
  theme?: string;
  [key: string]: unknown;
}

export default function TeamsTab() {
  // Add custom font styling
  useEffect(() => {
    // Apply font and font-smoothing styles
    const style = document.createElement('style');
    style.textContent = `
      * {
        font-family: 'Courier New', Courier, monospace !important;
        -webkit-font-smoothing: grayscale !important;
        -moz-osx-font-smoothing: grayscale !important;
        font-smoothing: grayscale !important;
        color: rgb(253, 246, 242) !important;
        background-color: #111827 !important;
      }
      button {
        margin: 1rem;
        padding: 1rem;
        cursor: pointer;
        font-weight: 600;
        font-size: 1rem;
        align-items: center !important;
        background-color: rgb(51, 72, 96) !important;
        background-image: none !important;
        border-bottom-color: rgb(39, 39, 42) !important;
        border-bottom-style: solid !important;
        border-bottom-width: 1px !important;
        border-collapse: collapse !important;
        border-image-outset: 0 !important;
        border-image-repeat: stretch !important;
        border-image-slice: 100% !important;
        border-image-source: none !important;
        border-image-width: 1 !important;
        border-left-color: rgb(39, 39, 42) !important;
        border-left-style: solid !important;
        border-left-width: 1px !important;
        border-right-color: rgb(39, 39, 42) !important;
        border-right-style: solid !important;
        border-right-width: 1px !important;
        border-top-color: rgb(39, 39, 42) !important;
        border-top-style: solid !important;
        border-top-width: 1px !important;
        box-sizing: border-box !important;
        border-radius: 0.5rem;
        color: rgb(253, 246, 242) !important;
      }
      button:hover {
        background-color: rgb(204, 132, 96) !important;
      }
    `;
    document.head.appendChild(style);

    // Cleanup on unmount
    return () => {
      document.head.removeChild(style);
    };
  }, []);

  const [isInitialized, setIsInitialized] = useState(false);
  const [authToken, setAuthToken] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [context, setContext] = useState<unknown>(null);
  const [tokenExchangeStatus, setTokenExchangeStatus] = useState<string>('idle');
  const [userInfo, setUserInfo] = useState<UserInfo | null>(null);
  const [teamsTestResult, setTeamsTestResult] = useState<string>('');
  const [teamsContextData, setTeamsContextData] = useState<Record<string, unknown> | string | null>(null);
  const [teamsContextCard, setTeamsContextCard] = useState<HTMLElement | null>(null);
  const [apiResultsCard, setApiResultsCard] = useState<HTMLElement | null>(null);

  // Helper function to create Teams Context Adaptive Card
  const createTeamsContextCard = (contextData: TeamsContext) => {
    const card = new AdaptiveCards.AdaptiveCard();

    // Add header
    const headerBlock = new AdaptiveCards.TextBlock();
    headerBlock.text = "Teams Context";
    headerBlock.size = AdaptiveCards.TextSize.Large;
    headerBlock.weight = AdaptiveCards.TextWeight.Bolder;
    card.addItem(headerBlock);

    // Add user information if available
    if (contextData.user) {
      const userHeader = new AdaptiveCards.TextBlock();
      userHeader.text = "User Information";
      userHeader.size = AdaptiveCards.TextSize.Medium;
      userHeader.weight = AdaptiveCards.TextWeight.Bolder;
      card.addItem(userHeader);

      if (contextData.user.id) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `User ID: ${contextData.user.id}`;
        card.addItem(fact);
      }
      if (contextData.user.displayName) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Display Name: ${contextData.user.displayName}`;
        card.addItem(fact);
      }
      if (contextData.user.userPrincipalName) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `UPN: ${contextData.user.userPrincipalName}`;
        card.addItem(fact);
      }
      if (contextData.user.email) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Email: ${contextData.user.email}`;
        card.addItem(fact);
      }
    }

    // Add team information if available
    if (contextData.team) {
      const teamHeader = new AdaptiveCards.TextBlock();
      teamHeader.text = "Team Information";
      teamHeader.size = AdaptiveCards.TextSize.Medium;
      teamHeader.weight = AdaptiveCards.TextWeight.Bolder;
      card.addItem(teamHeader);

      if (contextData.team.id) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Team ID: ${contextData.team.id}`;
        card.addItem(fact);
      }
      if (contextData.team.displayName) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Team Name: ${contextData.team.displayName}`;
        card.addItem(fact);
      }
      if (contextData.team.internalId) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Internal ID: ${contextData.team.internalId}`;
        card.addItem(fact);
      }
    }

    // Add channel information if available
    if (contextData.channel) {
      const channelHeader = new AdaptiveCards.TextBlock();
      channelHeader.text = "Channel Information";
      channelHeader.size = AdaptiveCards.TextSize.Medium;
      channelHeader.weight = AdaptiveCards.TextWeight.Bolder;
      card.addItem(channelHeader);

      if (contextData.channel.id) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Channel ID: ${contextData.channel.id}`;
        card.addItem(fact);
      }
      if (contextData.channel.displayName) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Channel Name: ${contextData.channel.displayName}`;
        card.addItem(fact);
      }
    }

    // Add app information if available
    if (contextData.app) {
      const appHeader = new AdaptiveCards.TextBlock();
      appHeader.text = "App Information";
      appHeader.size = AdaptiveCards.TextSize.Medium;
      appHeader.weight = AdaptiveCards.TextWeight.Bolder;
      card.addItem(appHeader);

      if (contextData.app.id) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `App ID: ${contextData.app.id}`;
        card.addItem(fact);
      }
      if (contextData.app.sessionId) {
        const fact = new AdaptiveCards.TextBlock();
        fact.text = `Session ID: ${contextData.app.sessionId}`;
        card.addItem(fact);
      }
    }

    // Add context type information
    const contextHeader = new AdaptiveCards.TextBlock();
    contextHeader.text = "Context Details";
    contextHeader.size = AdaptiveCards.TextSize.Medium;
    contextHeader.weight = AdaptiveCards.TextWeight.Bolder;
    card.addItem(contextHeader);

    const typeBlock = new AdaptiveCards.TextBlock();
    typeBlock.text = `Context Type: ${typeof contextData}`;
    card.addItem(typeBlock);

    const localeBlock = new AdaptiveCards.TextBlock();
    localeBlock.text = `Locale: ${contextData.locale || "Not available"}`;
    card.addItem(localeBlock);

    const themeBlock = new AdaptiveCards.TextBlock();
    themeBlock.text = `Theme: ${contextData.theme || "Not available"}`;
    card.addItem(themeBlock);

    return card.render();
  };

  // Helper function to create API Test Results Adaptive Card
  const createApiResultsCard = (status: string, userInfo: UserInfo | null, error: string | null) => {
    const card = new AdaptiveCards.AdaptiveCard();

    // Add header
    const headerBlock = new AdaptiveCards.TextBlock();
    headerBlock.text = "API Test Results";
    headerBlock.size = AdaptiveCards.TextSize.Large;
    headerBlock.weight = AdaptiveCards.TextWeight.Bolder;
    card.addItem(headerBlock);

    // Add status with color coding
    const statusColor = status === 'success' ? AdaptiveCards.TextColor.Good :
                       status === 'error' ? AdaptiveCards.TextColor.Attention :
                       status === 'exchanging' ? AdaptiveCards.TextColor.Warning : AdaptiveCards.TextColor.Default;

    const statusBlock = new AdaptiveCards.TextBlock();
    statusBlock.text = `Status: ${status.toUpperCase()}`;
    statusBlock.color = statusColor;
    statusBlock.weight = AdaptiveCards.TextWeight.Bolder;
    statusBlock.size = AdaptiveCards.TextSize.Medium;
    card.addItem(statusBlock);

    // Add user information if available
    if (userInfo) {
      const userHeader = new AdaptiveCards.TextBlock();
      userHeader.text = "User Information";
      userHeader.size = AdaptiveCards.TextSize.Medium;
      userHeader.weight = AdaptiveCards.TextWeight.Bolder;
      card.addItem(userHeader);

      const nameBlock = new AdaptiveCards.TextBlock();
      nameBlock.text = `Name: ${userInfo.name}`;
      card.addItem(nameBlock);

      const emailBlock = new AdaptiveCards.TextBlock();
      emailBlock.text = `Email: ${userInfo.email}`;
      card.addItem(emailBlock);

      const upnBlock = new AdaptiveCards.TextBlock();
      upnBlock.text = `UPN: ${userInfo.upn}`;
      card.addItem(upnBlock);

      const subBlock = new AdaptiveCards.TextBlock();
      subBlock.text = `Subject: ${userInfo.sub}`;
      card.addItem(subBlock);
    }

    // Add error information if available
    if (error) {
      const errorHeader = new AdaptiveCards.TextBlock();
      errorHeader.text = "Error Details";
      errorHeader.size = AdaptiveCards.TextSize.Medium;
      errorHeader.weight = AdaptiveCards.TextWeight.Bolder;
      errorHeader.color = AdaptiveCards.TextColor.Attention;
      card.addItem(errorHeader);

      const errorBlock = new AdaptiveCards.TextBlock();
      errorBlock.text = error;
      errorBlock.color = AdaptiveCards.TextColor.Attention;
      errorBlock.wrap = true;
      card.addItem(errorBlock);
    }

    return card.render();
  };

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
      const userData = typeof result === 'string' ? JSON.parse(result) as { user: UserInfo } : result;
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
      setTeamsTestResult('Testing Teams SDK functions...');
      
      // Test getting context again
      const currentContext = await microsoftTeams.app.getContext();
      console.log("Current context:", currentContext);
      
      // Update UI with results
      setTeamsContextData(typeof currentContext === 'string' ? currentContext : { ...currentContext });
      setTeamsTestResult('Teams SDK test completed successfully!');
      
      // Create and store the Adaptive Card
      const contextCard = createTeamsContextCard(currentContext as unknown as TeamsContext);
      setTeamsContextCard(contextCard || null);

    } catch (error) {
      console.error("Teams function test error:", error);
      setTeamsTestResult(`Teams SDK test failed: ${error instanceof Error ? error.message : 'Unknown error'}`);
      setTeamsContextData(null);
      setTeamsContextCard(null);
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

            // Create and store the Adaptive Card
            const resultsCard = createApiResultsCard('success', data.user as UserInfo, null);
            setApiResultsCard(resultsCard || null);
          } else {
            setTokenExchangeStatus('error');
            // Display debug information in the error message
            const debugInfo = data.debug ? `\n\nDebug Info:\n${Object.entries(data.debug).map(([key, value]) => `${key}: ${value}`).join('\n')}` : '';
            const errorMessage = `API Test Failed: ${data.error}${debugInfo}`;
            setError(errorMessage);

            // Create and store the Adaptive Card with error
            const resultsCard = createApiResultsCard('error', null, errorMessage);
            setApiResultsCard(resultsCard || null);
          }
        } else {
          const text = await response.text();
          const errorDetails = `API returned non-JSON: ${contentType || 'none'} - Response: ${text.substring(0, 100)}...`;
          setTokenExchangeStatus('error');
          setError(`API Test Failed: ${errorDetails}`);

          // Create and store the Adaptive Card with error
          const resultsCard = createApiResultsCard('error', null, `API Test Failed: ${errorDetails}`);
          setApiResultsCard(resultsCard || null);
        }
      } else {
        const text = await response.text();
        const errorDetails = `HTTP ${response.status} - Content-Type: ${contentType || 'none'} - Response: ${text.substring(0, 100)}...`;
        setTokenExchangeStatus('error');
        setError(`API Test Failed: ${errorDetails}`);

        // Create and store the Adaptive Card with error
        const resultsCard = createApiResultsCard('error', null, `API Test Failed: ${errorDetails}`);
        setApiResultsCard(resultsCard || null);
      }

    } catch (error) {
      setTokenExchangeStatus('error');
      const errorMessage = `API Test Error: ${error instanceof Error ? error.message : 'Unknown error'}`;
      setError(errorMessage);

      // Create and store the Adaptive Card with error
      const resultsCard = createApiResultsCard('error', null, errorMessage);
      setApiResultsCard(resultsCard || null);
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

        {/* Display Teams Context Adaptive Card */}
        {teamsContextCard && (
          <div className="p-4 bg-blue-50 rounded-lg">
            <h2 className="font-semibold text-blue-800 mb-2">Teams Context (Adaptive Card)</h2>
            <div ref={(el) => {
              if (el && teamsContextCard) {
                el.innerHTML = '';
                el.appendChild(teamsContextCard);
              }
            }} />
          </div>
        )}

        {teamsTestResult && (
          <div className="p-4 bg-blue-50 rounded-lg">
            <h2 className="font-semibold text-blue-800 mb-2">Teams SDK Test Results</h2>
            <div className="text-sm text-gray-700">
              <div className="mb-2">
                <strong>Status:</strong> {teamsTestResult}
              </div>
              {teamsContextData && (
                <div>
                  <strong>Teams Context Data (Raw JSON):</strong>
                  <div className="bg-gray-100 p-2 rounded text-xs font-mono break-all max-h-40 overflow-y-auto">
                    {typeof teamsContextData === 'object' && teamsContextData !== null ? (
                      <pre>{JSON.stringify(teamsContextData, null, 2)}</pre>
                    ) : typeof teamsContextData === 'string' ? (
                      <pre>{teamsContextData}</pre>
                    ) : null}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        <button
          onClick={testApiEndpoint}
          className="w-full bg-yellow-500 hover:bg-yellow-600 text-white font-medium py-2 px-4 rounded-lg transition-colors"
        >
          Test API Endpoint
        </button>

        {/* Display API Results Adaptive Card */}
        {apiResultsCard && (
          <div className="p-4 bg-yellow-50 rounded-lg">
            <h2 className="font-semibold text-yellow-800 mb-2">API Test Results (Adaptive Card)</h2>
            <div ref={(el) => {
              if (el && apiResultsCard) {
                el.innerHTML = '';
                el.appendChild(apiResultsCard);
              }
            }} />
          </div>
        )}

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

