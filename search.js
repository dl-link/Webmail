// Importing dependencies
import { userAccount, graphConfig, searchResults } from './main.js';

// Function to search through emails
function searchEmails(query) {
    // Clear previous search results
    searchResults.length = 0;

    // Fetch emails from Microsoft Graph API
    fetch(`https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages?$search="${query}"`, {
        headers: {
            'Authorization': 'Bearer ' + userAccount.accessToken
        }
    })
    .then(response => response.json())
    .then(data => {
        // Store the search results
        searchResults.push(...data.value);

        // Display success message
        document.getElementById('searchSuccess').style.display = 'block';
        document.getElementById('searchFailure').style.display = 'none';
    })
    .catch(error => {
        // Display failure message
        document.getElementById('searchFailure').style.display = 'block';
        document.getElementById('searchSuccess').style.display = 'none';

        console.error('Error:', error);
    });
}

// Event listener for search box
document.getElementById('searchBox').addEventListener('input', event => {
    // Search emails when user types in the search box
    searchEmails(event.target.value);
});

export { searchEmails };