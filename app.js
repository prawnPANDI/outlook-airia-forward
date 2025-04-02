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
                    
                    // API endpoint and configuration
                    const apiEndpoint = 'https://prodaus.api.airia.ai/v1/PipelineExecution/bc8e5a90-c46b-41a3-a0f6-72364ebf7a8f';
                    
                    // Send the email content to the API
                    fetch(apiEndpoint, {
                        method: 'POST',
                        headers: {
                            'Content-Type': 'application/json',
                            'X-API-KEY': 'ak-NjQ3NzIxNjk0fDE3NDM1ODE0Nzc2MzV8QWlyaWF8MXw2ODMwOTY0MTMg',
                            'Accept': 'application/json'
                        },
                        mode: 'cors',
                        credentials: 'omit',
                        body: JSON.stringify({
                            userInput: emailContent,
                            asyncOutput: false
                        })
                    })
                    .then(response => {
                        if (!response.ok) {
                            throw new Error(`HTTP error! status: ${response.status}`);
                        }
                        return response.json();
                    })
                    .then(data => {
                        showStatus('Email content sent successfully!');
                        console.log('API Response:', data);
                    })
                    .catch(error => {
                        console.error('Error details:', error);
                        showStatus('Error sending email content: ' + error.message);
                    });
                } else {
                    showStatus('Error getting email content: ' + result.error.message);
                }
            });
        } catch (error) {
            console.error('Error details:', error);
            showStatus('Error: ' + error.message);
        }
    });

    function showStatus(message) {
        statusElement.textContent = message;
    }
}); 