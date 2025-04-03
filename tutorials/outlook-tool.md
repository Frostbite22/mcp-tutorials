# Building an Outlook Integration Tool with MCP

This tutorial will guide you through creating an MCP tool that interfaces with Microsoft Outlook, allowing AI models to read emails, send messages, and manage calendar events.

## Prerequisites

- Node.js (v14+) installed
- Basic understanding of JavaScript and JSON-RPC
- Microsoft Azure account for registering an application
- Microsoft Graph API knowledge (basic)

## Project Setup

1. Create a new directory for your project:

```bash
mkdir mcp-outlook-tool
cd mcp-outlook-tool
```

2. Initialize a new Node.js project:

```bash
npm init -y
```

3. Install the required dependencies:

```bash
npm install express @microsoft/microsoft-graph-client isomorphic-fetch azure-identity
```

## Setting up Microsoft Graph API Access

Before we start coding, you need to register an application with Microsoft to get the necessary credentials:

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to "Azure Active Directory" > "App Registrations" > "New Registration"
3. Name your app (e.g., "MCP Outlook Tool")
4. Set the redirect URI to `http://localhost:3000/auth/callback`
5. Register the application
6. Note your Application (client) ID and Directory (tenant) ID
7. Under "Certificates & secrets", create a new client secret and note its value
8. Add Microsoft Graph API permissions:
   - Go to "API Permissions" > "Add a permission" > "Microsoft Graph" > "Delegated permissions"
   - Add the following permissions:
     - Mail.Read
     - Mail.Send
     - Calendars.ReadWrite
     - User.Read

## Creating the MCP Server

Create a new file called `server.js` with the following content:

```javascript
const express = require('express');
const { Client } = require('@microsoft/microsoft-graph-client');
const { ClientSecretCredential } = require('@azure/identity');
require('isomorphic-fetch');

const app = express();

// Microsoft Graph API credentials - replace with your own values
const config = {
  clientId: 'YOUR_CLIENT_ID',
  clientSecret: 'YOUR_CLIENT_SECRET',
  tenantId: 'YOUR_TENANT_ID',
  redirectUri: 'http://localhost:3000/auth/callback'
};

// Auth token storage - in production, use a proper database
let authTokens = {};

// Middleware to parse JSON requests
app.use(express.json());

// Create authenticated Microsoft Graph client
function getGraphClient(userId) {
  if (!authTokens[userId]) {
    throw new Error('User not authenticated');
  }
  
  const credential = new ClientSecretCredential(
    config.tenantId,
    config.clientId,
    config.clientSecret
  );
  
  return Client.init({
    authProvider: async (done) => {
      try {
        const token = authTokens[userId];
        done(null, token);
      } catch (error) {
        done(error, null);
      }
    }
  });
}

// Handle JSON-RPC requests
app.post('/', async (req, res) => {
  const { jsonrpc, method, params, id } = req.body;
  
  // Verify JSON-RPC 2.0 request
  if (jsonrpc !== '2.0' || !method || !id) {
    return res.json({
      jsonrpc: '2.0',
      error: { code: -32600, message: 'Invalid request' },
      id: null
    });
  }
  
  // Handle method calls
  try {
    let result;
    
    switch (method) {
      case 'initialize':
        // Respond to initialization with capabilities
        result = {
          protocolVersion: '2024-11-05',
          methods: {
            getEmails: {
              description: 'Get emails from a user\'s inbox',
              parameters: {
                type: 'object',
                properties: {
                  userId: {
                    type: 'string',
                    description: 'User ID for authentication'
                  },
                  folder: {
                    type: 'string',
                    description: 'Folder to fetch emails from (default: inbox)',
                    default: 'inbox'
                  },
                  count: {
                    type: 'number',
                    description: 'Number of emails to fetch (default: 10)',
                    default: 10
                  }
                },
                required: ['userId']
              }
            },
            sendEmail: {
              description: 'Send an email',
              parameters: {
                type: 'object',
                properties: {
                  userId: {
                    type: 'string',
                    description: 'User ID for authentication'
                  },
                  to: {
                    type: 'array',
                    items: { type: 'string' },
                    description: 'Recipients email addresses'
                  },
                  subject: {
                    type: 'string',
                    description: 'Email subject'
                  },
                  body: {
                    type: 'string',
                    description: 'Email body (HTML supported)'
                  },
                  cc: {
                    type: 'array',
                    items: { type: 'string' },
                    description: 'CC recipients email addresses'
                  },
                  bcc: {
                    type: 'array',
                    items: { type: 'string' },
                    description: 'BCC recipients email addresses'
                  }
                },
                required: ['userId', 'to', 'subject', 'body']
              }
            },
            createCalendarEvent: {
              description: 'Create a calendar event',
              parameters: {
                type: 'object',
                properties: {
                  userId: {
                    type: 'string',
                    description: 'User ID for authentication'
                  },
                  subject: {
                    type: 'string',
                    description: 'Event subject/title'
                  },
                  start: {
                    type: 'string',
                    description: 'Start date and time (ISO format)'
                  },
                  end: {
                    type: 'string',
                    description: 'End date and time (ISO format)'
                  },
                  location: {
                    type: 'string',
                    description: 'Event location'
                  },
                  attendees: {
                    type: 'array',
                    items: { type: 'string' },
                    description: 'Email addresses of attendees'
                  },
                  body: {
                    type: 'string',
                    description: 'Event description'
                  }
                },
                required: ['userId', 'subject', 'start', 'end']
              }
            }
          }
        };
        break;
        
      case 'getEmails':
        // Check for required parameters
        if (!params || !params.userId) {
          throw { code: -32602, message: 'Invalid params - userId is required' };
        }
        
        try {
          const client = getGraphClient(params.userId);
          const folder = params.folder || 'inbox';
          const count = params.count || 10;
          
          const emailsResponse = await client
            .api(`/me/mailFolders/${folder}/messages`)
            .top(count)
            .select('id,subject,bodyPreview,receivedDateTime,from,toRecipients,hasAttachments')
            .orderBy('receivedDateTime DESC')
            .get();
          
          result = {
            emails: emailsResponse.value.map(email => ({
              id: email.id,
              subject: email.subject,
              preview: email.bodyPreview,
              received: email.receivedDateTime,
              from: email.from.emailAddress,
              to: email.toRecipients.map(r => r.emailAddress),
              hasAttachments: email.hasAttachments
            }))
          };
        } catch (error) {
          throw { code: -32603, message: `Graph API error: ${error.message}` };
        }
        break;
        
      case 'sendEmail':
        // Check for required parameters
        if (!params || !params.userId || !params.to || !params.subject || !params.body) {
          throw { code: -32602, message: 'Invalid params - userId, to, subject, and body are required' };
        }
        
        try {
          const client = getGraphClient(params.userId);
          
          const emailMessage = {
            subject: params.subject,
            body: {
              contentType: 'HTML',
              content: params.body
            },
            toRecipients: params.to.map(email => ({
              emailAddress: { address: email }
            }))
          };
          
          // Add CC recipients if provided
          if (params.cc && params.cc.length > 0) {
            emailMessage.ccRecipients = params.cc.map(email => ({
              emailAddress: { address: email }
            }));
          }
          
          // Add BCC recipients if provided
          if (params.bcc && params.bcc.length > 0) {
            emailMessage.bccRecipients = params.bcc.map(email => ({
              emailAddress: { address: email }
            }));
          }
          
          await client.api('/me/sendMail').post({
            message: emailMessage,
            saveToSentItems: true
          });
          
          result = { success: true, message: 'Email sent successfully' };
        } catch (error) {
          throw { code: -32603, message: `Graph API error: ${error.message}` };
        }
        break;
        
      case 'createCalendarEvent':
        // Check for required parameters
        if (!params || !params.userId || !params.subject || !params.start || !params.end) {
          throw { code: -32602, message: 'Invalid params - userId, subject, start, and end are required' };
        }
        
        try {
          const client = getGraphClient(params.userId);
          
          const event = {
            subject: params.subject,
            start: {
              dateTime: new Date(params.start).toISOString(),
              timeZone: 'UTC'
            },
            end: {
              dateTime: new Date(params.end).toISOString(),
              timeZone: 'UTC'
            }
          };
          
          // Add location if provided
          if (params.location) {
            event.location = {
              displayName: params.location
            };
          }
          
          // Add body/description if provided
          if (params.body) {
            event.body = {
              contentType: 'HTML',
              content: params.body
            };
          }
          
          // Add attendees if provided
          if (params.attendees && params.attendees.length > 0) {
            event.attendees = params.attendees.map(email => ({
              emailAddress: {
                address: email
              },
              type: 'required'
            }));
          }
          
          const createdEvent = await client.api('/me/events').post(event);
          
          result = {
            success: true,
            eventId: createdEvent.id,
            webLink: createdEvent.webLink
          };
        } catch (error) {
          throw { code: -32603, message: `Graph API error: ${error.message}` };
        }
        break;
        
      default:
        // Handle unknown methods
        throw { code: -32601, message: `Method not found: ${method}` };
    }
    
    // Return successful response
    return res.json({
      jsonrpc: '2.0',
      result,
      id
    });
  } catch (error) {
    // Handle errors
    console.error('Error processing request:', error);
    
    // Format error for JSON-RPC response
    const errorResponse = {
      jsonrpc: '2.0',
      error: {
        code: error.code || -32603,
        message: error.message || 'Internal error'
      },
      id
    };
    
    return res.json(errorResponse);
  }
});

// Authentication route - this is a simplified implementation
// In production, use proper OAuth flow with PKCE
app.get('/auth', (req, res) => {
  const authUrl = `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/authorize?client_id=${config.clientId}&response_type=code&redirect_uri=${encodeURIComponent(config.redirectUri)}&response_mode=query&scope=User.Read%20Mail.Read%20Mail.Send%20Calendars.ReadWrite`;
  res.redirect(authUrl);
});

app.get('/auth/callback', async (req, res) => {
  const code = req.query.code;
  if (!code) {
    return res.status(400).send('Authorization code missing');
  }
  
  try {
    // In a real implementation, exchange code for tokens here
    // For this tutorial, we'll use a placeholder
    const userId = 'user123'; // In production, get real user ID
    authTokens[userId] = 'placeholder_token'; // In production, store the actual token
    
    res.send(`
      <h1>Authentication Successful</h1>
      <p>Your userId for MCP calls is: <strong>${userId}</strong></p>
      <p>Use this ID in your MCP requests</p>
    `);
  } catch (error) {
    console.error('Auth error:', error);
    res.status(500).send('Authentication failed: ' + error.message);
  }
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`MCP Outlook Tool server running on port ${PORT}`);
  console.log(`To authenticate, visit: http://localhost:${PORT}/auth`);
});
```

## Running the Server

1. Replace the placeholders in the config object with your actual Azure application credentials:
   - `YOUR_CLIENT_ID`
   - `YOUR_CLIENT_SECRET`
   - `YOUR_TENANT_ID`

2. Start the server:

```bash
node server.js
```

3. Authenticate by visiting `http://localhost:3000/auth` in your browser

## Testing the Outlook Tool

You can test your Outlook tool using curl. Remember to use the userId returned during authentication:

### Get Emails

```bash
curl -X POST http://localhost:3000 \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "method": "getEmails",
    "params": {
      "userId": "user123",
      "folder": "inbox",
      "count": 5
    },
    "id": 1
  }'
```

### Send Email

```bash
curl -X POST http://localhost:3000 \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "method": "sendEmail",
    "params": {
      "userId": "user123",
      "to": ["recipient@example.com"],
      "subject": "Test from MCP Outlook Tool",
      "body": "<p>This is a test email sent from the MCP Outlook integration.</p>"
    },
    "id": 2
  }'
```

### Create Calendar Event

```bash
curl -X POST http://localhost:3000 \
  -H "Content-Type: application/json" \
  -d '{
    "jsonrpc": "2.0",
    "method": "createCalendarEvent",
    "params": {
      "userId": "user123",
      "subject": "Test Meeting",
      "start": "2025-04-10T14:00:00Z",
      "end": "2025-04-10T15:00:00Z",
      "location": "Conference Room A",
      "attendees": ["colleague@example.com"],
      "body": "<p>Quarterly review meeting</p>"
    },
    "id": 3
  }'
```

## Security Considerations

This tutorial provides a simplified implementation. For production use, consider:

1. **Proper OAuth Flow**: Implement a complete OAuth 2.0 flow with PKCE
2. **Token Management**: Store and refresh tokens securely
3. **Rate Limiting**: Add protection against excessive requests
4. **Input Validation**: Add more robust validation for all inputs
5. **Logging**: Add structured logging with sensitive data redacted
6. **Error Handling**: Improve error handling with appropriate status codes

## Extending the Tool

To enhance your Outlook integration tool, consider:

1. **Email Search**: Add functionality to search emails by keyword
2. **Attachment Handling**: Allow downloading or viewing attachments
3. **Meeting Management**: Add functionality to update/delete meetings
4. **Contact Management**: Add functionality to access and manage contacts
5. **Folder Management**: Add support for creating/managing email folders

## Complete Code

The complete code for this tutorial is available in the [examples/outlook-tool](../examples/outlook-tool/) directory.
