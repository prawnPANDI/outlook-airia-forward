Office.onReady(() => {
    // Ensure we're in a supported version of Office
    if (!Office.context.mailbox) {
        showStatus('This add-in requires Outlook.');
        return;
    }

    // Get the button element
    const sendButton = document.getElementById('sendButton');
    const statusElement = document.getElementById('status');

    // Create log display
    const logDisplay = document.createElement('div');
    logDisplay.style.cssText = `
        background-color: #f5f5f5;
        padding: 15px;
        border-radius: 5px;
        margin: 10px 0;
        max-height: 300px;
        overflow-y: auto;
        font-family: monospace;
        font-size: 12px;
        white-space: pre-wrap;
    `;
    statusElement.parentNode.insertBefore(logDisplay, statusElement);

    // Function to add log entry
    function addLogEntry(message, type = 'info') {
        const timestamp = new Date().toLocaleTimeString();
        const entry = document.createElement('div');
        entry.style.cssText = `
            margin: 5px 0;
            padding: 5px;
            border-left: 3px solid ${type === 'error' ? '#ff4444' : '#4CAF50'};
            background-color: ${type === 'error' ? '#fff5f5' : '#f5fff5'};
        `;
        entry.textContent = `[${timestamp}] ${message}`;
        logDisplay.appendChild(entry);
        logDisplay.scrollTop = logDisplay.scrollHeight;
        console.log(`[${timestamp}] ${message}`);
    }

    // Function to format email content with all available information
    function formatEmailContent(item) {
        return new Promise((resolve, reject) => {
            // Get the email body
            item.body.getAsync(Office.CoercionType.Text, (result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const emailInfo = {
                        subject: item.subject || '',
                        sender: item.from ? item.from.emailAddress : '',
                        recipients: {
                            to: item.to ? item.to.map(r => r.emailAddress).join(', ') : '',
                            cc: item.cc ? item.cc.map(r => r.emailAddress).join(', ') : '',
                            bcc: item.bcc ? item.bcc.map(r => r.emailAddress).join(', ') : ''
                        },
                        receivedDateTime: item.dateTimeCreated ? new Date(item.dateTimeCreated).toISOString() : '',
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
            // Clear previous logs
            logDisplay.innerHTML = '';
            addLogEntry('Starting email processing...');
            
            // Get the current email item
            const item = Office.context.mailbox.item;
            addLogEntry('Step 1: Getting email item...');
            
            // Format the email content
            addLogEntry('Step 2: Reading email content...');
            const emailContent = await formatEmailContent(item);
            addLogEntry('Step 3: Email content read successfully');
            
            // Display the formatted JSON
            addLogEntry('Step 4: Preparing data preview...');
            displayJSON(emailContent);
            addLogEntry('Step 5: Data preview ready');
            
            addLogEntry('Step 6: Preparing API request...');
            
            // API endpoint and configuration
            const apiEndpoint = 'https://prodaus.api.airia.ai/v1/PipelineExecution/bc8e5a90-c46b-41a3-a0f6-72364ebf7a8f/';
            const requestData = {
                userId: "019540ef-47e9-7b57-a89b-2c521617064f",
                userInput: JSON.stringify(emailContent, null, 2)
            };
            
            addLogEntry(`API Request Data: ${JSON.stringify(requestData, null, 2)}`);
            
            // Send the email content to the API
            addLogEntry('Step 7: Sending data to API...');
            try {
                const response = await fetch(apiEndpoint, {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'X-API-KEY': 'ak-NjQ3NzIxNjk0fDE3NDM1ODE0Nzc2MzV8QWlyaWF8MXw2ODMwOTY0MTMg',
                        'Accept': 'application/json',
                        'Origin': window.location.origin
                    },
                    mode: 'cors',
                    credentials: 'omit',
                    body: JSON.stringify(requestData)
                });

                if (!response.ok) {
                    const errorText = await response.text();
                    addLogEntry(`API Error Response: ${errorText}`, 'error');
                    throw new Error(`HTTP error! status: ${response.status}, message: ${errorText}`);
                }

                addLogEntry('Step 8: API request successful, processing response...');
                const data = await response.json();
                addLogEntry('Step 9: Complete! Email content sent successfully!');
                addLogEntry(`API Response: ${JSON.stringify(data, null, 2)}`);
                console.log('API Response:', data);
            } catch (fetchError) {
                console.error('Fetch Error:', fetchError);
                if (fetchError.name === 'TypeError' && fetchError.message.includes('Failed to fetch')) {
                    addLogEntry('CORS Error: The API endpoint is not accessible from this domain. This might be due to CORS restrictions.', 'error');
                    addLogEntry('Technical Details: ' + fetchError.message, 'error');
                } else {
                    addLogEntry('API Error: ' + fetchError.message, 'error');
                }
                throw fetchError;
            }
        } catch (error) {
            console.error('Error details:', error);
            addLogEntry('Error: ' + error.message, 'error');
        }
    });

    function showStatus(message) {
        statusElement.textContent = message;
        addLogEntry(message);
    }
}); 