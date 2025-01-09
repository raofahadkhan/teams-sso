// import {
//     PublicClientApplication,
//     Configuration,
//     LogLevel,
//   } from "@azure/msal-browser";

//   const msalConfig: Configuration = {
//     auth: {
//       clientId: "4ea3481a-86f1-4730-8d75-0c5e2f621d9b", // Your Client ID
//       authority:
//         "https://login.microsoftonline.com/5d83f397-271d-40b5-8f97-b400080e94a5", // Your Tenant ID (Authority)
//       redirectUri:
//         "https://eb23-103-74-22-42.ngrok-free.app/auth-end", // Ensure this is registered in Azure AD
//       navigateToLoginRequestUrl: false,
//     },
//     cache: {
//       cacheLocation: "sessionStorage",
//       storeAuthStateInCookie: false,
//     },
//     system: {
//       allowRedirectInIframe: true, // Allow authentication flows inside an iframe (important for Teams)
//       loggerOptions: {
//         loggerCallback: (level, message) => {
//           console.log(message); // Verbose logging for debugging
//         },
//         logLevel: LogLevel.Verbose,
//       },
//     },
//   };

//   const msalInstance = new PublicClientApplication(msalConfig);

//   // Async function to initialize MSAL
//   export const initializeMsal = async () => {
//     try {
//       await msalInstance.initialize();
//       console.log("MSAL initialized successfully");
//     } catch (error) {
//       console.error("Error initializing MSAL", error);
//     }
//   };

//   export default msalInstance;

import {
  PublicClientApplication,
  Configuration,
  LogLevel,
} from "@azure/msal-browser";

const msalConfig: Configuration = {
  auth: {
    clientId: "4ea3481a-86f1-4730-8d75-0c5e2f621d9b", // Replace with your App ID
    authority:
      "https://login.microsoftonline.com/5d83f397-271d-40b5-8f97-b400080e94a5", // Multi-tenant authority
    redirectUri: "https://eb23-103-74-22-42.ngrok-free.app/auth-end", // Ensure this is registered in Azure AD
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
  system: {
    allowRedirectInIframe: true, // Important for Teams authentication
    loggerOptions: {
      loggerCallback: (level, message) => {
        console.log(message); // Debug logging
      },
      logLevel: LogLevel.Verbose,
    },
  },
};

const msalInstance = new PublicClientApplication(msalConfig);

// Initialize MSAL
export const initializeMsal = async () => {
  try {
    await msalInstance.initialize();
    console.log("MSAL initialized successfully");
  } catch (error) {
    console.error("Error initializing MSAL:", error);
  }
};

export default msalInstance;
