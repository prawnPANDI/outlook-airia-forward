# Outlook Email Content Sender Add-in

This is a simple Outlook add-in that allows you to send the content of the currently displayed email to an API endpoint.

## Features

- Sends email content (subject, body, sender, recipients, and timestamp) to a specified API endpoint
- Simple and clean user interface
- Status feedback for successful/failed operations

## Setup Instructions

1. Replace the API endpoint in `app.js`:
   - Find the line with `const apiEndpoint = 'https://your-api-endpoint.com/process-email';`
   - Replace it with your actual API endpoint

2. Deploy the add-in:
   - You'll need to host these files on a web server (HTTPS required)
   - Update the `SourceLocation` in `manifest.xml` to point to your hosted location

3. Sideload the add-in in Outlook:
   - For Windows: [Sideload Office Add-ins in Outlook](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing)
   - For Mac: [Sideload Office Add-ins on Mac](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/sideload-outlook-add-ins-for-testing#sideload-an-add-in-on-mac)

## Usage

1. Open an email in Outlook
2. Click the add-in button in the add-in bar
3. Click the "Send to API" button to send the email content to your API endpoint
4. Check the status message to confirm the operation was successful

## Development

To test locally:

1. Set up a local development server (e.g., using Python's `http.server` or Node.js's `http-server`)
2. Update the `SourceLocation` in `manifest.xml` to point to your local server
3. Sideload the add-in as described above

## Security Notes

- Ensure your API endpoint is secured with appropriate authentication
- Consider implementing rate limiting on your API endpoint
- Handle sensitive email content appropriately

## Requirements

- Outlook 2016 or later
- Modern web browser
- HTTPS hosting for the add-in files 