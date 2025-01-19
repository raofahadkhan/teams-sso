"use client";

import { useEffect } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { msalInstance, initializeMsal, codeVerifier } from "../teams-tab/authconfig";

const AuthEnd: React.FC = () => {
  useEffect(() => {
    console.log("[AuthEnd] Initializing Teams...");

    // Use the callback form of initialize to ensure it's done before handling redirect
    microsoftTeams.initialize(() => {
      console.log("[AuthEnd] Teams initialization complete. Handling redirect...");
      handleRedirect();
    });

    const handleRedirect = async (): Promise<void> => {
      try {
        // Initialize MSAL
        await initializeMsal();
        console.log("[AuthEnd] MSAL initialization successful.");

        // Handle MSAL redirect
        const response = await msalInstance.handleRedirectPromise();

        if (response) {
          console.log("[AuthEnd] Auth response received:", response);

          // Check if codeVerifier is available
          if (!codeVerifier) {
            throw new Error("Code verifier is missing. Unable to complete token exchange.");
          }

          const tokenRequest = {
            code: response.code!,
            scopes: ["User.Read", "api://your-api-id/access_as_user"],
            redirectUri: "https://teams-sso.vercel.app/auth-end",
            codeVerifier, // Use the codeVerifier
          };

          const tokenResponse = await msalInstance.acquireTokenByCode(tokenRequest);
          console.log("[AuthEnd] Token acquired:", tokenResponse);

          const accountDetails = {
            name: tokenResponse.account?.name || "N/A",
            username: tokenResponse.account?.username || "N/A",
            homeAccountId: tokenResponse.account?.homeAccountId || "N/A",
          };

          console.log("[AuthEnd] Extracted account details:", accountDetails);

          microsoftTeams.authentication.notifySuccess(JSON.stringify(accountDetails));
          console.log("[AuthEnd] Authentication succeeded, closing window...");
          window.close();
        } else {
          console.warn("[AuthEnd] No redirect response. Possible misconfiguration or silent auth failure.");
          microsoftTeams.authentication.notifyFailure("No response found");
        }
      } catch (err) {
        if (err instanceof Error) {
          console.error("[AuthEnd] Error encountered during redirect handling:", err);
          microsoftTeams.authentication.notifyFailure(err.message || "Redirect error");
        } else {
          console.error("[AuthEnd] Unknown error encountered:", err);
          microsoftTeams.authentication.notifyFailure("An unknown error occurred");
        }
      }
    };
  }, []);

  return <div style={{ color: "white" }}>Processing authentication...</div>;
};

export default AuthEnd;
