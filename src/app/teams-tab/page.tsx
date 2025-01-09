// "use client";

// import { useEffect, useState } from "react";
// import * as microsoftTeams from "@microsoft/teams-js";
// import msalInstance, { initializeMsal } from "./authconfig";

// const TeamsAuth = () => {
//   const [authState, setAuthState] = useState<any>({
//     isAuthenticated: false,
//     error: null,
//     loading: true,
//     user: null,
//   });

//   useEffect(() => {
//     const authenticateWithTeams = async () => {
//       try {
//         // Initialize Teams SDK
//         microsoftTeams.initialize(() => {
//           console.log("Teams SDK initialized");
//         });

//         // Initialize MSAL
//         await initializeMsal();
//         console.log("MSAL initialized");

//         // Define authentication request
//         const request = {
//           scopes: [
//             "User.Read",
//             "api://4ea3481a-86f1-4730-8d75-0c5e2f621d9b/access_as_user",
//           ],
//         };

//         // Silent authentication
//         const response = await msalInstance.ssoSilent(request);
//         console.log("Silent authentication success:", response);

//         setAuthState({
//           isAuthenticated: true,
//           user: response.account,
//           error: null,
//           loading: false,
//         });
//       } catch (error:any) {
//         console.error("Silent authentication failed:", error);

//         if (error.name === "InteractionRequiredAuthError") {
//           // If interaction is required, use Teams authentication popup
//           microsoftTeams.authentication.authenticate({
//             url: `https://login.microsoftonline.com/5d83f397-271d-40b5-8f97-b400080e94a5/oauth2/v2.0/authorize?client_id=4ea3481a-86f1-4730-8d75-0c5e2f621d9b&response_type=code&redirect_uri=https://eb23-103-74-22-42.ngrok-free.app/auth-end&scope=openid email profile offline_access User.Read`,
//             width: 600,
//             height: 535,
//             successCallback: (result:any) => {
//               console.log("Popup authentication success:", result);
//               setAuthState({
//                 isAuthenticated: true,
//                 user: result,
//                 error: null,
//                 loading: false,
//               });
//             },
//             failureCallback: (popupError:any) => {
//               console.error("Popup authentication failed:", popupError);
//               setAuthState({
//                 isAuthenticated: false,
//                 error: "Authentication failed",
//                 loading: false,
//                 user: null,
//               });
//             },
//           });
//         } else {
//           setAuthState({
//             isAuthenticated: false,
//             error: "Authentication failed",
//             loading: false,
//             user: null,
//           });
//         }
//       }
//     };

//     authenticateWithTeams();
//   }, []);

//   if (authState.loading) {
//     return <div style={{ color: "white" }}>Authenticating...</div>;
//   }

//   if (authState.error) {
//     return (
//       <div style={{ color: "white" }}>
//         <h2>Authentication failed</h2>
//         <p>{authState.error}</p>
//       </div>
//     );
//   }

//   if (authState.isAuthenticated) {
//     const email =
//       authState.user?.idTokenClaims?.email ||
//       authState.user?.idTokenClaims?.preferred_username ||
//       "No email available";

//     return (
//       <div style={{ color: "white" }}>
//         <h1>Welcome, {authState.user?.name}!</h1>
//         <p>Your email: {email}</p>
//         <p>You have successfully authenticated with Teams SSO.</p>
//       </div>
//     );
//   }

//   return <div style={{ color: "white" }}>Authentication required</div>;
// };

// export default TeamsAuth;
"use client";

import { useEffect, useState } from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import msalInstance, { initializeMsal } from "./authconfig";

const TeamsAuth = () => {
  const [authState, setAuthState] = useState<any>({
    isAuthenticated: false,
    error: null,
    loading: true,
    user: null,
    email: null,
  });

  useEffect(() => {
    const authenticateWithTeams = async () => {
      try {
        // Initialize Teams SDK
        microsoftTeams.initialize(() => {
          console.log("Teams SDK initialized");
        });

        // Get Teams Context (Fallback for User Info)
        microsoftTeams.getContext((context) => {
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

        // Define authentication request
        const request = {
          scopes: [
            "User.Read",
            "api://4ea3481a-86f1-4730-8d75-0c5e2f621d9b/access_as_user",
          ],
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
          microsoftTeams.authentication.authenticate({
            url: `https://login.microsoftonline.com/5d83f397-271d-40b5-8f97-b400080e94a5/oauth2/v2.0/authorize?client_id=4ea3481a-86f1-4730-8d75-0c5e2f621d9b&response_type=code&redirect_uri=https://eb23-103-74-22-42.ngrok-free.app/auth-end&scope=openid email profile offline_access User.Read`,
            width: 600,
            height: 535,
            successCallback: (result:any) => {
              console.log("Popup authentication success:", result);

              setAuthState({
                isAuthenticated: true,
                user: result,
                email: result.idTokenClaims?.email || result.idTokenClaims?.preferred_username || null,
                error: null,
                loading: false,
              });

              console.log("Authenticated User Email from Popup:", result.idTokenClaims?.email);
            },
            failureCallback: (popupError:any) => {
              console.error("Popup authentication failed:", popupError);
              setAuthState((prevState:any) => ({
                ...prevState,
                isAuthenticated: false,
                error: "Authentication failed",
                loading: false,
              }));
            },
          });
        } else {
          // Silent auth failed with other errors
          setAuthState((prevState:any) => ({
            ...prevState,
            isAuthenticated: false,
            error: error.message || "Authentication failed",
            loading: false,
          }));
        }
      }
    };

    authenticateWithTeams();
  }, []);

  if (authState.loading) {
    return <div style={{ color: "white" }}>Authenticating...</div>;
  }

  if (authState.error) {
    console.warn("Authentication failed. Attempting to log email...");
    if (authState.email) {
      console.log("Fallback User Email:", authState.email);
    }
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
