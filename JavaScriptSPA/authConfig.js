const msalConfig = {
    auth: {
      clientId: "",
      authority: "https://login.microsoftonline.com/",
      redirectUri: "http://localhost:4200",
    },
    cache: {
      cacheLocation: "sessionStorage", // This configures where your cache will be stored
      storeAuthStateInCookie: false, // Set this to "true" if you're having issues on Internet Explorer 11 or Edge
    }
  };

  // Add scopes for the ID token to be used at Microsoft identity platform endpoints.
  const loginRequest = {
    scopes: ["openid", "profile", "User.Read"]
  };

  // Add scopes for the access token to be used at Microsoft Graph API endpoints.
  const tokenRequest = {
    scopes: ["Mail.Read"]
  };

  const silentRequest = {
    scopes: ["openid", "profile", "User.Read", "Mail.Read"]
};

const logoutRequest = {}