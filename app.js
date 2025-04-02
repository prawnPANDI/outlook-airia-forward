Office.onReady(() => {
    // Ensure we're in a supported version of Office
    if (!Office.context.mailbox) {
        showStatus('This add-in requires Outlook.');
        return;
    }

    // Get the button element
    const sendButton = document.getElementById('sendButton');
    const statusElement = document.getElementById('status');

    // Function to format email content with all available information
    function formatEmailContent(item) {
        const emailInfo = {
            subject: item.subject || '',
            sender: item.from ? item.from.emailAddress : '',
            recipients: {
                to: item.to ? item.to.map(r => r.emailAddress).join(', ') : '',
                cc: item.cc ? item.cc.map(r => r.emailAddress).join(', ') : '',
                bcc: item.bcc ? item.bcc.map(r => r.emailAddress).join(', ') : ''
            },
            receivedDateTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : '',
            categories: item.categories ? item.categories.join(', ') : '',
            importance: item.importance || '',
            hasAttachments: item.hasAttachments || false,
            attachments: item.attachments ? item.attachments.map(att => att.name).join(', ') : '',
            conversationId: item.conversationId || '',
            internetMessageId: item.internetMessageId || '',
            body: ''
        };

        // Get the email body
        return new Promise((resolve, reject) => {
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    emailInfo.body = result.value;
                    resolve(emailInfo);
                } else {
                    reject(new Error(result.error.message));
                }
            });
        });
    }

    // Add click event handler
    sendButton.addEventListener('click', async () => {
        try {
            // Get the current email item
            const item = Office.context.mailbox.item;
            
            // Format the email content
            const emailContent = await formatEmailContent(item);
            
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
                    userId: "019540ef-47e9-7b57-a89b-2c521617064f",
                    userInput: JSON.stringify(emailContent, null, 2)
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
        } catch (error) {
            console.error('Error details:', error);
            showStatus('Error: ' + error.message);
        }
    });

    function showStatus(message) {
        statusElement.textContent = message;
    }
}); 