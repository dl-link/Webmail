// Importing necessary modules
import { msalConfig, userAccount, graphConfig } from './auth.js';
import { fetchEmails, displayEmail } from './mail.js';
import { searchEmails } from './search.js';

// DOM Elements
const loginButton = document.getElementById('loginButton');
const logoutButton = document.getElementById('logoutButton');
const composeButton = document.getElementById('composeButton');
const sendButton = document.getElementById('sendButton');
const searchBox = document.getElementById('searchBox');
const emailList = document.getElementById('emailList');
const emailContent = document.getElementById('emailContent');

// MSAL.js UserAgentApplication
const msalApplication = new Msal.UserAgentApplication(msalConfig);

// Event listeners
loginButton.addEventListener('click', login);
logoutButton.addEventListener('click', logout);
composeButton.addEventListener('click', composeEmail);
sendButton.addEventListener('click', sendEmail);
searchBox.addEventListener('input', searchEmails);

// Login function
function login() {
  msalApplication.loginPopup(graphConfig).then(function (loginResponse) {
    // Login Success
    userAccount = msalApplication.getAccount();
    console.log('loginSuccess', userAccount);
    fetchEmails();
  }).catch(function (error) {
    // Login Failure
    console.log('loginFailure', error);
  });
}

// Logout function
function logout() {
  msalApplication.logout();
}

// Compose Email function
function composeEmail() {
  // Code to compose email goes here
}

// Send Email function
function sendEmail() {
  // Code to send email goes here
}

// Fetch Emails function
function fetchEmails() {
  // Code to fetch emails goes here
}

// Display Email function
function displayEmail(emailId) {
  // Code to display email goes here
}

// Search Emails function
function searchEmails(searchQuery) {
  // Code to search emails goes here
}