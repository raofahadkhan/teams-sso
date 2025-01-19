import {
  PublicClientApplication,
  Configuration,
  LogLevel,
} from "@azure/msal-browser";

let codeVerifier: string | null = null; // Store the codeVerifier globally for later use

const generatePKCECodes = async () => {
  const generateRandomString = (length: number): string => {
    const charset =
      "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-._~";
    let randomString = "";
    const randomValues = new Uint8Array(length);
    window.crypto.getRandomValues(randomValues);
    for (let i = 0; i < randomValues.length; i++) {
      randomString += charset[randomValues[i] % charset.length];
    }
    return randomString;
  };

  const sha256 = async (plain: string): Promise<string> => {
    const encoder = new TextEncoder();
    const data = encoder.encode(plain);
    const hash = await crypto.subtle.digest("SHA-256", data);

    // Fix: Convert Uint8Array to a regular array before using String.fromCharCode
    const hashArray = Array.from(new Uint8Array(hash));
    return btoa(
      String.fromCharCode(...hashArray) // Spread the array to convert to characters
    )
      .replace(/\+/g, "-")
      .replace(/\//g, "_")
      .replace(/=+$/, "");
  };

  const codeVerifier = generateRandomString(128);
  const codeChallenge = await sha256(codeVerifier);
  return { codeVerifier, codeChallenge };
};

const msalConfig: Configuration = {
  auth: {
    clientId: "5572abc7-7a99-448a-9f62-134da3f27e9e", // Replace with your App ID
    authority: "https://login.microsoftonline.com/common", // Multi-tenant authority
    redirectUri: "https://teams-sso.vercel.app/auth-end", // Ensure this matches Azure AD settings
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

export const initializeMsal = async () => {
  try {
    await msalInstance.initialize();
    console.log("MSAL initialized successfully");
  } catch (error) {
    console.error("Error initializing MSAL:", error);
  }
};

// Expose PKCE generation and the codeVerifier
export { msalInstance, generatePKCECodes, codeVerifier };
