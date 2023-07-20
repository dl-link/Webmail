Shared Dependencies:

1. Exported Variables:
   - `msalConfig`: Configuration object for MSAL.js.
   - `userAccount`: User account object after successful authentication.
   - `graphConfig`: Configuration object for Microsoft Graph API.
   - `mailList`: Array to store the fetched emails.
   - `searchResults`: Array to store the search results.

2. Data Schemas:
   - `Email`: Schema for an email object, including properties like sender, recipient, subject, body, etc.
   - `User`: Schema for a user object, including properties like id, name, email, etc.

3. ID Names of DOM Elements:
   - `loginButton`: Button for user login.
   - `logoutButton`: Button for user logout.
   - `composeButton`: Button to compose a new email.
   - `sendButton`: Button to send the composed email.
   - `searchBox`: Input field for searching emails.
   - `emailList`: Container to display the list of emails.
   - `emailContent`: Container to display the content of a selected email.

4. Message Names:
   - `loginSuccess`: Message when user login is successful.
   - `loginFailure`: Message when user login fails.
   - `sendSuccess`: Message when email is sent successfully.
   - `sendFailure`: Message when email sending fails.
   - `searchSuccess`: Message when search is successful.
   - `searchFailure`: Message when search fails.

5. Function Names:
   - `login()`: Function to handle user login.
   - `logout()`: Function to handle user logout.
   - `composeEmail()`: Function to compose a new email.
   - `sendEmail()`: Function to send the composed email.
   - `searchEmails()`: Function to search through emails.
   - `fetchEmails()`: Function to fetch emails from Microsoft Exchange Online.
   - `displayEmail()`: Function to display the content of a selected email.