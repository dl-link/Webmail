// Importing necessary dependencies
import { userAccount, graphConfig } from './auth.js';

// Array to store fetched emails
let mailList = [];

// Function to fetch emails from Microsoft Exchange Online
async function fetchEmails() {
    const accessToken = await acquireTokenSilent({
        scopes: graphConfig.mailReadScopes,
        account: userAccount
    });

    const response = await fetch(graphConfig.graphMeEndpoint + 'mailFolders/inbox/messages', {
        headers: {
            Authorization: 'Bearer ' + accessToken
        }
    });

    if (response.ok) {
        const data = await response.json();
        mailList = data.value;
        displayEmails();
    } else {
        throw new Error(response.status);
    }
}

// Function to display the list of emails
function displayEmails() {
    const emailList = document.getElementById('emailList');
    emailList.innerHTML = '';

    mailList.forEach((email, index) => {
        const emailItem = document.createElement('li');
        emailItem.textContent = email.subject;
        emailItem.addEventListener('click', () => displayEmail(index));
        emailList.appendChild(emailItem);
    });
}

// Function to display the content of a selected email
function displayEmail(index) {
    const emailContent = document.getElementById('emailContent');
    const email = mailList[index];

    emailContent.innerHTML = `
        <h2>${email.subject}</h2>
        <p>From: ${email.from.emailAddress.address}</p>
        <p>To: ${email.toRecipients.map(recipient => recipient.emailAddress.address).join(', ')}</p>
        <p>${email.body.content}</p>
    `;
}

// Function to compose a new email
function composeEmail() {
    const emailContent = document.getElementById('emailContent');
    emailContent.innerHTML = `
        <h2>New Email</h2>
        <form id="composeForm">
            <label for="to">To:</label>
            <input type="email" id="to" required>
            <label for="subject">Subject:</label>
            <input type="text" id="subject" required>
            <label for="body">Body:</label>
            <textarea id="body" required></textarea>
            <button type="submit" id="sendButton">Send</button>
        </form>
    `;

    document.getElementById('composeForm').addEventListener('submit', sendEmail);
}

// Function to send the composed email
async function sendEmail(event) {
    event.preventDefault();

    const to = document.getElementById('to').value;
    const subject = document.getElementById('subject').value;
    const body = document.getElementById('body').value;

    const email = {
        message: {
            subject: subject,
            body: {
                contentType: 'Text',
                content: body
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: to
                    }
                }
            ]
        },
        saveToSentItems: 'true'
    };

    const accessToken = await acquireTokenSilent({
        scopes: graphConfig.mailSendScopes,
        account: userAccount
    });

    const response = await fetch(graphConfig.graphMeEndpoint + 'sendMail', {
        method: 'POST',
        headers: {
            Authorization: 'Bearer ' + accessToken,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(email)
    });

    if (response.ok) {
        alert('Email sent successfully!');
        fetchEmails();
    } else {
        throw new Error(response.status);
    }
}

export { fetchEmails, composeEmail };