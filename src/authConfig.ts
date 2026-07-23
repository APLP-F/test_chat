const isGitHubPages =
  window.location.hostname === "aplp-f.github.io";

const redirectUri = isGitHubPages
  ? "https://aplp-f.github.io/test_chat/auth.html"
  : "http://localhost:5173/auth.html";

export const msalConfig = {
  auth: {
    clientId: "84ef3bdb-a4d5-4094-8e8f-c763b1a3edbb",
    authority:
      "https://login.microsoftonline.com/436efeb8-20ec-47a9-9dbb-3deab23539f5",
    redirectUri,
  },

  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

export const copilotLoginRequest = {
  scopes: [
    "https://api.powerplatform.com/CopilotStudio.Copilots.Invoke",
  ],
};