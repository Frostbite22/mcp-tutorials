# Outlook MCP Tool - Python Implementation Tutorial

This tutorial guides you through implementing and using an MCP (Model Context Protocol) tool for interacting with Microsoft Outlook using the Microsoft Graph API.

## Overview

The Outlook MCP Tool allows AI models and applications to:
- Read emails from an Outlook inbox
- Send emails through Outlook
- Create calendar events
- List upcoming appointments

## Prerequisites

- Python 3.8+
- Microsoft Azure account with registered application
- Microsoft 365 account with appropriate permissions

## Setup and Installation

### 1. Install Required Dependencies

```bash
pip install -r requirements.txt
```

The requirements.txt file should include:
- Flask
- flask-cors
- requests
- msal (Microsoft Authentication Library)

### 2. Azure Application Registration

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Click "New registration"
4. Provide a name for your application
5. Select appropriate account type (typically "Single tenant")
6. Click "Register"

### 3. Configure API Permissions

1. In your registered app, go to "API permissions"
2. Click "Add a permission"
3. Select "Microsoft Graph"
4. Choose "Application permissions"
5. Add the following permissions:
   - Mail.Read
   - Mail.Send
   - Calendars.ReadWrite
6. Click "Add permissions"
7. Grant admin consent for these permissions

### 4. Create Client Secret

1. In your app, go to "Certificates & secrets"
2. Click "New client secret"
3. Add a description and select expiration
4. Copy the generated secret value (visible only once)

### 5. Configure the Tool

Create a `config.json` file in your project directory:

```json
{
  "client_id": "YOUR_CLIENT_ID",
  "client_secret": "YOUR_CLIENT_SECRET",
  "tenant_id": "YOUR_TENANT_ID",
  "authority": "https://login.microsoftonline.com/YOUR_TENANT_ID",
  "scope": ["https://graph.microsoft.com/.default"],
  "endpoint": "https://graph.microsoft.com/v1.0",
  "user_email": "optional-specific-user@yourdomain.com"
}
```

## Implementation Details

### Tool Structure

The Outlook MCP Tool consists of:
- Authentication module using MSAL
- Graph API request handler
- MCP method implementations
- Flask server for JSON-RPC API

### Core Components

#### Authentication

```python
def get_token():
    """
    Get access token using MSAL
    """
    app = msal.ConfidentialClientApplication(
        config["client_id"],
        authority=config["authority"],
        client_credential=config["client_secret"]
    )
    
    result = app.acquire_token_for_client(scopes=config["scope"])
    
    if "access_token" in result:
        return result["access_token"]
    else:
        logger.error(f"Authentication error: {result.get('error')}, {result.get('error_description')}")
        return None
```

#### Graph API Requests

```python
def make_graph_request(method, endpoint, token, data=None):
    """
    Make a request to Microsoft Graph API
    """
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }
    
    url = f"{config['endpoint']}/{endpoint}"
    
    if method.lower() == 'get':
        response = requests.get(url, headers=headers)
    elif method.lower() == 'post':
        response = requests.post(url, headers=headers, json=data)
    else:
        raise ValueError(f"Unsupported method: {method}")
    
    if response.status_code >= 400:
        logger.error(f"Graph API error: {response.status_code} - {response.text}")
        return None
    
    return response.json()
```

## MCP Method Implementations

### 1. Reading Emails

```python
def read_emails(params):
    """
    Read emails from inbox
    
    Params:
    - count: Number of emails to retrieve (default: 10)
    - filter: Filter string for emails (optional)
    
    Returns:
    - List of email objects
    """
    token = get_token()
    if not token:
        return {"error": "Authentication failed"}
    
    count = params.get("count", 10)
    filter_str = params.get("filter", "")
    
    endpoint = f"users/{config.get('user_email', 'me')}/messages"
    if filter_str:
        endpoint += f"?$filter={filter_str}"
    endpoint += f"&$top={count}&$orderby=receivedDateTime desc"
    
    response = make_graph_request("get", endpoint, token)
    if not response:
        return {"error": "Failed to retrieve emails"}
    
    emails = []
    for email in response.get("value", []):
        emails.append({
            "id": email.get("id"),
            "subject": email.get("subject"),
            "sender": email.get("sender", {}).get("emailAddress", {}).get("address"),
            "received": email.get("receivedDateTime"),
            "body": email.get("bodyPreview"),
            "isRead": email.get("isRead")
        })
    
    return {"emails": emails}
```

### 2. Sending Emails

```python
def send_email(params):
    """
    Send an email
    
    Params:
    - to: Email address of recipient
    - subject: Email subject
    - body: Email body
    - cc: CC recipients (optional)
    - bcc: BCC recipients (optional)
    
    Returns:
    - Success or error message
    """
    token = get_token()
    if not token:
        return {"error": "Authentication failed"}
    
    # Required parameters
    to = params.get("to")
    subject = params.get("subject")
    body = params.get("body")
    
    if not all([to, subject, body]):
        return {"error": "Missing required parameters: to, subject, body"}
    
    # Optional parameters
    cc = params.get("cc", [])
    bcc = params.get("bcc", [])
    
    # Format recipients
    to_recipients = [{"emailAddress": {"address": email}} for email in to.split(",")]
    cc_recipients = [{"emailAddress": {"address": email}} for email in cc] if cc else []
    bcc_recipients = [{"emailAddress": {"address": email}} for email in bcc] if bcc else []
    
    # Create message
    message = {
        "message": {
            "subject": subject,
            "body": {
                "contentType": "HTML",
                "content": body
            },
            "toRecipients": to_recipients,
            "ccRecipients": cc_recipients,
            "bccRecipients": bcc_recipients
        },
        "saveToSentItems": "true"
    }
    
    endpoint = f"users/{config.get('user_email', 'me')}/sendMail"
    response = make_graph_request("post", endpoint, token, message)
    
    if response is None:  # No response means success for this endpoint
        return {"success": True, "message": "Email sent successfully"}
    else:
        return {"error": "Failed to send email", "details": response}
```

### 3. Creating Calendar Events

```python
def create_calendar_event(params):
    """
    Create a calendar event
    
    Params:
    - subject: Event subject/title
    - start: Start time (ISO format)
    - end: End time (ISO format)
    - body: Event description (optional)
    - location: Event location (optional)
    - attendees: List of attendee email addresses (optional)
    
    Returns:
    - Created event details or error
    """
    token = get_token()
    if not token:
        return {"error": "Authentication failed"}
    
    # Required parameters
    subject = params.get("subject")
    start_time = params.get("start")
    end_time = params.get("end")
    
    if not all([subject, start_time, end_time]):
        return {"error": "Missing required parameters: subject, start, end"}
    
    # Optional parameters
    body = params.get("body", "")
    location = params.get("location", "")
    attendees = params.get("attendees", [])
    
    # Format attendees
    attendee_list = []
    for email in attendees:
        attendee_list.append({
            "emailAddress": {
                "address": email
            },
            "type": "required"
        })
    
    # Create event
    event = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": body
        },
        "start": {
            "dateTime": start_time,
            "timeZone": "UTC"
        },
        "end": {
            "dateTime": end_time,
            "timeZone": "UTC"
        },
        "attendees": attendee_list
    }
    
    if location:
        event["location"] = {
            "displayName": location
        }
    
    endpoint = f"users/{config.get('user_email', 'me')}/calendar/events"
    response = make_graph_request("post", endpoint, token, event)
    
    if not response:
        return {"error": "Failed to create calendar event"}
    
    return {
        "success": True,
        "event": {
            "id": response.get("id"),
            "subject": response.get("subject"),
            "start": response.get("start", {}).get("dateTime"),
            "end": response.get("end", {}).get("dateTime"),
            "webLink": response.get("webLink")
        }
    }
```

### 4. Listing Appointments

```python
def list_appointments(params):
    """
    List upcoming appointments
    
    Params:
    - days: Number of days to look ahead (default: 7)
    - max_events: Maximum number of events to return (default: 10)
    
    Returns:
    - List of upcoming events
    """
    token = get_token()
    if not token:
        return {"error": "Authentication failed"}
    
    days = params.get("days", 7)
    max_events = params.get("max_events", 10)
    
    # Calculate date range
    now = datetime.datetime.utcnow().isoformat() + 'Z'
    end_date = (datetime.datetime.utcnow() + datetime.timedelta(days=days)).isoformat() + 'Z'
    
    # Build query
    endpoint = f"users/{config.get('user_email', 'me')}/calendarView?startDateTime={now}&endDateTime={end_date}&$top={max_events}&$orderby=start/dateTime"
    
    response = make_graph_request("get", endpoint, token)
    if not response:
        return {"error": "Failed to retrieve appointments"}
    
    events = []
    for event in response.get("value", []):
        events.append({
            "id": event.get("id"),
            "subject": event.get("subject"),
            "start": event.get("start", {}).get("dateTime"),
            "end": event.get("end", {}).get("dateTime"),
            "location": event.get("location", {}).get("displayName"),
            "organizer": event.get("organizer", {}).get("emailAddress", {}).get("address"),
            "isOnline": event.get("isOnlineMeeting", False)
        })
    
    return {"appointments": events}
```

## Running the MCP Server

The tool implements a Flask server that handles JSON-RPC requests:

```python
# MCP method mapping
MCP_METHODS = {
    "read_emails": read_emails,
    "send_email": send_email,
    "create_calendar_event": create_calendar_event,
    "list_appointments": list_appointments
}

@app.route('/mcp', methods=['POST'])
def handle_mcp_request():
    """
    Handle MCP JSON-RPC requests
    """
    try:
        data = request.json
        logger.info(f"Received MCP request: {data}")
        
        # Check for required fields
        if 'jsonrpc' not in data or data['jsonrpc'] != '2.0':
            return jsonify({"jsonrpc": "2.0", "error": {"code": -32600, "message": "Invalid Request"}, "id": None})
        
        if 'method' not in data:
            return jsonify({"jsonrpc": "2.0", "error": {"code": -32600, "message": "Method not specified"}, "id": data.get('id')})
        
        method = data['method']
        params = data.get('params', {})
        request_id = data.get('id')
        
        # Check if method exists
        if method not in MCP_METHODS:
            return jsonify({"jsonrpc": "2.0", "error": {"code": -32601, "message": f"Method '{method}' not found"}, "id": request_id})
        
        # Execute method
        result = MCP_METHODS[method](params)
        
        # Return response
        return jsonify({"jsonrpc": "2.0", "result": result, "id": request_id})
        
    except Exception as e:
        logger.error(f"Error processing request: {str(e)}")
        return jsonify({"jsonrpc": "2.0", "error": {"code": -32603, "message": f"Internal error: {str(e)}"}, "id": request.json.get('id') if hasattr(request, 'json') else None})
```

To run the server:

```python
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 3000))
    app.run(host="0.0.0.0", port=port, debug=False)
    logger.info(f"MCP server started on port {port}")
```

Start the server with:

```bash
python outlook_tool.py
```

## Example Usage

### Reading Emails

```json
{
  "jsonrpc": "2.0",
  "method": "read_emails",
  "params": {
    "count": 5,
    "filter": "isRead eq false"
  },
  "id": 1
}
```

### Sending an Email

```json
{
  "jsonrpc": "2.0",
  "method": "send_email",
  "params": {
    "to": "recipient@example.com",
    "subject": "Hello from MCP Tool",
    "body": "<p>This is a test email sent using the Outlook MCP Tool.</p>"
  },
  "id": 2
}
```

### Creating a Calendar Event

```json
{
  "jsonrpc": "2.0",
  "method": "create_calendar_event",
  "params": {
    "subject": "Project Meeting",
    "start": "2025-04-10T15:00:00",
    "end": "2025-04-10T16:00:00",
    "body": "<p>Discuss project progress and next steps</p>",
    "location": "Conference Room B",
    "attendees": ["colleague1@example.com", "colleague2@example.com"]
  },
  "id": 3
}
```

### Listing Appointments

```json
{
  "jsonrpc": "2.0",
  "method": "list_appointments",
  "params": {
    "days": 14,
    "max_events": 20
  },
  "id": 4
}
```

## Troubleshooting

### Common Issues

1. **Authentication Failures**
   - Verify your client ID, client secret, and tenant ID
   - Ensure proper permissions are granted and admin consent provided
   - Check token expiration and renewal logic

2. **Permission Errors**
   - Verify the application has been granted the necessary permissions
   - Ensure admin consent has been provided for these permissions

3. **Invalid Parameters**
   - Check the parameter format for each method
   - Ensure required parameters are provided

### Debugging

The tool includes logging that can help diagnose issues:

```python
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)
```

For more detailed logs, change the logging level to DEBUG:

```python
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
```

## Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [MSAL Python Documentation](https://docs.microsoft.com/en-us/azure/active-directory/develop/msal-python-samples)
- [MCP Protocol Specification](https://modelcontextprotocol.io/docs/)
- [Flask Documentation](https://flask.palletsprojects.com/)
- [JSON-RPC Specification](https://www.jsonrpc.org/specification)
