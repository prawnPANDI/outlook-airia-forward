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
        return new Promise((resolve, reject) => {
            // Get the email body
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    // Log the categories property to understand its structure
                    console.log('Categories property:', item.categories);
                    
                    let categoriesValue = '';
                    if (item.categories) {
                        if (Array.isArray(item.categories)) {
                            categoriesValue = item.categories.join(', ');
                        } else if (typeof item.categories === 'string') {
                            categoriesValue = item.categories;
                        } else {
                            categoriesValue = String(item.categories);
                        }
                    }

                    const emailInfo = {
                        subject: item.subject || '',
                        sender: item.from ? item.from.emailAddress : '',
                        recipients: {
                            to: item.to ? item.to.map(r => r.emailAddress).join(', ') : '',
                            cc: item.cc ? item.cc.map(r => r.emailAddress).join(', ') : '',
                            bcc: item.bcc ? item.bcc.map(r => r.emailAddress).join(', ') : ''
                        },
                        receivedDateTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : '',
                        categories: categoriesValue,
                        importance: item.importance || '',
                        hasAttachments: item.hasAttachments || false,
                        attachments: item.attachments ? item.attachments.map(att => att.name).join(', ') : '',
                        conversationId: item.conversationId || '',
                        internetMessageId: item.internetMessageId || '',
                        body: result.value
                    };
                    resolve(emailInfo);
                } else {
                    reject(new Error('Failed to get email body: ' + result.error.message));
                }
            });
        });
    }

    // Function to display JSON in a formatted way
    function displayJSON(jsonData) {
        const pre = document.createElement('pre');
        pre.style.cssText = `
            background-color: #f5f5f5;
            padding: 15px;
            border-radius: 5px;
            overflow-x: auto;
            font-family: monospace;
            font-size: 12px;
            margin: 10px 0;
            max-height: 300px;
            overflow-y: auto;
        `;
        pre.textContent = JSON.stringify(jsonData, null, 2);
        
        // Remove any existing pre element
        const existingPre = document.querySelector('pre');
        if (existingPre) {
            existingPre.remove();
        }
        
        // Insert the new pre element before the status element
        statusElement.parentNode.insertBefore(pre, statusElement);
    }

    // Add click event handler
    sendButton.addEventListener('click', async () => {
        try {
            showStatus('Loading email content...');
            
            // Get the current email item
            const item = Office.context.mailbox.item;
            
            // Format the email content
            const emailContent = await formatEmailContent(item);
            
            // Display the formatted JSON
            displayJSON(emailContent);
            
            showStatus('Sending to API...');
            
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