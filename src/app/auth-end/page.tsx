"use client";

import { useEffect } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import msalInstance, { initializeMsal } from "../teams-tab/authconfig";

const AuthEnd = () => {
  useEffect(() => {
    console.log("[AuthEnd] Initializing Teams...");

    // Use the callback form of initialize to ensure it's done before handling redirect
    microsoftTeams.initialize(() => {
      console.log(
        "[AuthEnd] Teams initialization complete. Handling redirect...",
      );
      handleRedirect();
    });

    const handleRedirect = async () => {
      try {
        // Initialize MSAL
        await initializeMsal();
        console.log("[AuthEnd] MSAL initialization successful.");

        // Handle MSAL redirect
        const response = await msalInstance.handleRedirectPromise();

        if (response) {
          console.log("[AuthEnd] Auth response received:", response);

          const accountDetails = {
            name: response.account?.name || "N/A",
            username: response.account?.username || "N/A",
            homeAccountId: response.account?.homeAccountId || "N/A",
          };

          console.log("[AuthEnd] Extracted account details:", accountDetails);

          const successMessage = JSON.stringify(accountDetails);

          // Teams SDK is initialized, now it's safe to notify
          microsoftTeams.authentication.notifySuccess(successMessage);
          console.log("[AuthEnd] Authentication succeeded, closing window...");
          window.close();
        } else {
          console.warn(
            "[AuthEnd] No redirect response. Possible misconfiguration or silent auth failure.",
          );
          microsoftTeams.authentication.notifyFailure("No response found");
        }
      } catch (err) {
        if (err instanceof Error) {
          console.error(
            "[AuthEnd] Error encountered during redirect handling:",
            err,
          );
          microsoftTeams.authentication.notifyFailure(
            err.message || "Redirect error",
          );
        } else {
          console.error("[AuthEnd] Unknown error encountered:", err);
          microsoftTeams.authentication.notifyFailure(
            "An unknown error occurred",
          );
        }
      }
    };
  }, []);

  return <div style={{ color: "white" }}>Processing authentication...</div>;
};

export default AuthEnd;