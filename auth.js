// Importing MSAL.js
import * as Msal from 'msal';

// Configuration for MSAL.js
export const msalConfig = {
    auth: {
        clientId: 'your-app-client-id', // Replace with your app client id
        authority: 'https://login.microsoftonline.com/your-tenant-id', // Replace with your tenant id
        redirectUri: 'http://localhost:3000', // Replace with your redirect uri
    },
    cache: {
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false,
    },
};

// Creating a new instance of MSAL.js
const msalInstance = new Msal.UserAgentApplication(msalConfig);

// User account object
export let userAccount = null;

// Function to handle user login
export function login() {
    const loginRequest = {
        scopes: ['openid', 'profile', 'User.Read', 'Mail.ReadWrite'],
    };

    msalInstance.loginPopup(loginRequest)
        .then((loginResponse) => {
            userAccount = msalInstance.getAccount();
            console.log('loginSuccess', `Logged in as ${userAccount.name}`);
        })
        .catch((error) => {
            console.log('loginFailure', error);
        });
}

// Function to handle user logout
export function logout() {
    msalInstance.logout();
}

// Function to get access token
export function getAccessToken(scopes) {
    const accessTokenRequest = {
        scopes,
        account: userAccount,
    };

    return msalInstance.acquireTokenSilent(accessTokenRequest)
        .catch((error) => {
            console.log('Failed to acquire token silently. Acquiring token using popup', error);
            return msalInstance.acquireTokenPopup(accessTokenRequest);
        });
}