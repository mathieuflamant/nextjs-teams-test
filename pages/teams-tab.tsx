import { useEffect } from 'react';
import * as microsoftTeams from "@microsoft/teams-js";

export default function TeamsTab() {
  useEffect(() => {
    microsoftTeams.app.initialize().then(() => {
      microsoftTeams.authentication.getAuthToken({
        successCallback: (token) => {
          fetch("/api/auth/teams", {
            method: "POST",
            headers: { Authorization: `Bearer ${token}` },
          });
        },
        failureCallback: (error) => {
          console.error("SSO error", error);
        },
      });
    });
  }, []);

  return <div>Welcome to Teams Test App</div>;
}

