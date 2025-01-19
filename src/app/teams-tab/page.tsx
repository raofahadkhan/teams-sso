"use client";

import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { msalInstance, initializeMsal, generatePKCECodes } from "./authconfig";

const TeamsAuth: React.FC = () => {
  const [authState, setAuthState] = useState<any>({
    isAuthenticated: false,
    error: null,
    loading: true,
    user: null,
    email: null,
  });

  useEffect(() => {
    const authenticateWithTeams = async (): Promise<void> => {
      try {
        // Initialize Teams SDK
        microsoftTeams.initialize(() => {
          console.log("Teams SDK initialized");
        });

        // Get Teams Context (Fallback for User Info)
        microsoftTeams.getContext((context: microsoftTeams.Context) => {
          console.log("Teams Context:", context);
          if (context && context.userPrincipalName) {
            console.log("User Email from Teams Context:", context.userPrincipalName);
            setAuthState((prevState:any) => ({
              ...prevState,
              email: context.userPrincipalName,
            }));
          }
        });

        // Initialize MSAL
        await initializeMsal();
        console.log("MSAL initialized");

        // Generate PKCE codes
        const { codeChallenge } = await generatePKCECodes();

        // Define authentication request
        const request = {
          scopes: [
            "User.Read",
            "api://5572abc7-7a99-448a-9f62-134da3f27e9e/access_as_user",
          ],
          codeChallenge,
          codeChallengeMethod: "S256",
        };

        // Attempt silent authentication
        const response = await msalInstance.ssoSilent(request);
        console.log("Silent authentication success:", response);

        const email =
          response.account?.idTokenClaims?.email ||
          response.account?.idTokenClaims?.preferred_username;

        setAuthState({
          isAuthenticated: true,
          user: response.account,
          email: email || null,
          error: null,
          loading: false,
        });

        console.log("Authenticated User Email:", email);
      } catch (error: any) {
        console.error("Silent authentication failed:", error);

        if (error.name === "InteractionRequiredAuthError") {
          // Fallback to interactive authentication
          const { codeChallenge } = await generatePKCECodes(); // Regenerate PKCE codes
          microsoftTeams.authentication.authenticate({
            url: `https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=5572abc7-7a99-448a-9f62-134da3f27e9e&response_type=code&redirect_uri=https://teams-sso.vercel.app/auth-end&scope=openid email profile offline_access User.Read&code_challenge=${codeChallenge}&code_challenge_method=S256`,
            width: 600,
            height: 535,
            successCallback: (result: string) => {
              console.log("Popup authentication success:", result);

              const parsedResult = JSON.parse(result); // Parse result if needed
              setAuthState({
                isAuthenticated: true,
                user: parsedResult,
                email: parsedResult?.idTokenClaims?.email || parsedResult?.idTokenClaims?.preferred_username || null,
                error: null,
                loading: false,
              });

              console.log("Authenticated User Email from Popup:", parsedResult?.idTokenClaims?.email);
            },
            failureCallback: (popupError: string) => {
              console.error("Popup authentication failed:", popupError);
              setAuthState({
                isAuthenticated: false,
                error: "Authentication failed",
                loading: false,
                user: null,
              });
            },
          });
        } else {
          setAuthState({
            isAuthenticated: false,
            error: error.message || "Authentication failed",
            loading: false,
            user: null,
          });
        }
      }
    };

    authenticateWithTeams();
  }, []);

  if (authState.loading) {
    return <div style={{ color: "white" }}>Authenticating...</div>;
  }

  if (authState.error) {
    return (
      <div style={{ color: "white" }}>
        <h2>Authentication failed</h2>
        <p>{authState.error}</p>
        {authState.email && <p>User Email: {authState.email}</p>}
      </div>
    );
  }

  if (authState.isAuthenticated) {
    const email = authState.email || "No email available";

    return (
      <div style={{ color: "white" }}>
        <h1>Welcome, {authState.user?.name || "User"}!</h1>
        <p>Your email: {email}</p>
        <p>You have successfully authenticated with Teams SSO.</p>
      </div>
    );
  }

  return <div style={{ color: "white" }}>Authentication required</div>;
};

export default TeamsAuth;
