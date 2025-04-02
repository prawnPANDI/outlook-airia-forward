Office.onReady(() => {
    // Ensure we're in a supported version of Office
    if (!Office.context.mailbox) {
        showStatus('This add-in requires Outlook.');
        return;
    }

    // Get the button element
    const sendButton = document.getElementById('sendButton');
    const statusElement = document.getElementById('status');

    // Add click event handler
    sendButton.addEventListener('click', async () => {
        try {
            // Get the current email item
            const item = Office.context.mailbox.item;
            
            // Get the email body
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const emailContent = result.value;
                    
                    // Replace this URL with your actual API endpoint
                    const apiEndpoint = 'https://your-api-endpoint.com/process-email';
                    
                    // Send the email content to the API
                    fetch(apiEndpoint, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                        },
                        body: JSON.stringify({
                            subject: item.subject,
                            body: emailContent,
                            sender: item.from.emailAddress,
                            recipients: item.to.map(recipient => recipient.emailAddress),
                            receivedDateTime: item.dateTimeCreated
                        })
                    })
                    .then(response => response.json())
                    .then(data => {
                        showStatus('Email content sent successfully!');
                    })
                    .catch(error => {
                        showStatus('Error sending email content: ' + error.message);
                    });
                } else {
                    showStatus('Error getting email content: ' + result.error.message);
                }
            });
        } catch (error) {
            showStatus('Error: ' + error.message);
        }
    });

    function showStatus(message) {
        statusElement.textContent = message;
    }
}); 