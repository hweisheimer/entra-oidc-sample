// Replace these values with your Entra ID app registration details.
// clientId: Application (client) ID from the Azure portal
// authority: https://login.microsoftonline.com/<your-tenant-id>
export const msalConfig = {
  auth: {
    clientId: "[CLIENT-ID]",
    authority: "https://login.microsoftonline.com/[TENANT-ID]",
    redirectUri: window.location.origin,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const loginRequest = {
  // scopes: ["openid", "profile", "email"],
  scopes: ["api://[CLIENT-ID]/impersonate_user"]
};

// web-api, running locally (see ../../web-api).
export const apiConfig = {
  profileEndpoint: "http://localhost:5100/api/profile",
};
