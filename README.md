# BambooHR Time Off Request Sync

A Google Apps Script solution for syncing time-off requests from BambooHR to Google Sheets with automatic data processing and formatting.

## Features

- **Time Off Request Sync**: Automatically syncs time-off requests from BambooHR
- **Multiple Versions**: Basic and enhanced versions with different features
- **Automatic Sheet Creation**: Creates dedicated sheets for data organization
- **Data Processing**: Calculates working days and formats data
- **Error Handling**: Comprehensive error checking and user notifications
- **Custom Menu**: Easy-to-use Google Sheets menu integration

## Prerequisites

- Google Apps Script project
- BambooHR account with API access
- Google Sheets for data storage
- Basic knowledge of JavaScript and Google Apps Script

## Setup Instructions

### 1. Google Apps Script Project Setup

1. Create a new Google Apps Script project
2. Copy the code from one of the files into your project:
   - `Code2.js` - Basic version with calculated days
   - `Code3.js` - Enhanced version with formatted headers and additional fields
   - `Untitled.js` - Simple version with basic functionality

### 2. BambooHR API Setup

1. Log into your BambooHR account
2. Navigate to Settings > API
3. Generate an API key with appropriate permissions
4. Note your BambooHR subdomain (found in your URL)

### 3. Configuration

Replace the placeholder values in your script:
```javascript
// Configuration - Replace with your actual values
var apiKey = 'YOUR_BAMBOOHR_API_KEY_HERE';
var subdomain = 'YOUR_BAMBOOHR_SUBDOMAIN_HERE';
```

## Usage

### Basic Usage

1. Open your Google Sheet
2. You'll see a new "BambooHR" menu in the toolbar
3. Click "Sync Time Off Requests"
4. The script will fetch and display time-off requests

### Data Output

The script creates a spreadsheet with the following columns:

#### Basic Version (Code2.js, Untitled.js)
- Employee ID
- Name
- Start Date
- End Date
- Status
- Type
- Reason
- Hours/Days
- Calculated Days (Code2.js only)

#### Enhanced Version (Code3.js)
- Request Created Date
- Employee ID
- Name
- Start Date
- End Date
- Status
- Type
- Reason
- Amount
- Unit

## API Functions

### Core Functions

- `onOpen()`: Creates the BambooHR menu in Google Sheets
- `syncTimeOffRequests()`: Main function to sync time-off data
- `calculateDays(startDate, endDate)`: Calculates working days between dates

### Error Handling

The script includes comprehensive error handling:
- API key validation
- Response format checking
- Network error handling
- User-friendly error messages

## Configuration Options

### Date Range

The script automatically fetches data from one month ago to today:
```javascript
var today = new Date();
var oneMonthAgo = new Date();
oneMonthAgo.setMonth(today.getMonth() - 1);
```

### Working Day Calculation

The script calculates working days with custom weights:
- Weekdays: 1.25 days
- Weekends: 1.69 days

This can be modified in the `calculateDays()` function.

## Example Use Cases

### 1. HR Reporting

```javascript
// Generate monthly time-off reports
function generateMonthlyReport() {
  syncTimeOffRequests();
  
  // Additional processing
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Filter by status
  var approvedRequests = data.filter(row => row[4] === 'approved');
  
  console.log(`Total requests: ${data.length}`);
  console.log(`Approved requests: ${approvedRequests.length}`);
}
```

### 2. Leave Balance Tracking

```javascript
// Track leave balances by employee
function trackLeaveBalances() {
  syncTimeOffRequests();
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Group by employee
  var employeeBalances = {};
  
  data.forEach(row => {
    var employeeId = row[1];
    var days = parseFloat(row[8]) || 0;
    
    if (!employeeBalances[employeeId]) {
      employeeBalances[employeeId] = 0;
    }
    
    employeeBalances[employeeId] += days;
  });
  
  console.log(employeeBalances);
}
```

### 3. Automated Notifications

```javascript
// Send notifications for pending requests
function checkPendingRequests() {
  syncTimeOffRequests();
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  var pendingRequests = data.filter(row => row[4] === 'pending');
  
  if (pendingRequests.length > 0) {
    // Send email notification
    var message = `You have ${pendingRequests.length} pending time-off requests to review.`;
    // MailApp.sendEmail('hr@company.com', 'Pending Time-Off Requests', message);
  }
}
```

## Data Structure

### Time Off Request Object

```javascript
{
  "employeeId": "123",
  "name": "John Doe",
  "start": "2024-01-15",
  "end": "2024-01-17",
  "status": {
    "status": "approved"
  },
  "type": {
    "name": "Vacation"
  },
  "notes": {
    "employee": "Family vacation"
  },
  "amount": {
    "amount": 3,
    "unit": "days"
  },
  "created": "2024-01-10T10:00:00Z"
}
```

## Configuration

### API Limits

- **Rate Limits**: Check your BambooHR plan for specific limits
- **Authentication**: API key required for all requests
- **Data Format**: JSON for all requests and responses

### Error Codes

Common error responses:
- `401`: Invalid API key
- `403`: Insufficient permissions
- `404`: Resource not found
- `429`: Rate limit exceeded
- `500`: Server error

## Security Considerations

- **API Keys**: Never commit API keys to version control
- **Data Privacy**: Ensure compliance with data protection regulations
- **Access Control**: Restrict script access to authorized users
- **Data Handling**: Securely process and store employee information

## Best Practices

### 1. API Key Management

```javascript
// Store API key in script properties
function getApiKey() {
  return PropertiesService.getScriptProperties().getProperty('BAMBOOHR_API_KEY');
}

function getSubdomain() {
  return PropertiesService.getScriptProperties().getProperty('BAMBOOHR_SUBDOMAIN');
}
```

### 2. Error Handling

```javascript
try {
  var response = UrlFetchApp.fetch(url, options);
  var data = JSON.parse(response.getContentText());
  
  if (!data || data.length === 0) {
    console.log('No data found');
    return;
  }
  
  // Process data
} catch (e) {
  console.error('API Error:', e.message);
  SpreadsheetApp.getUi().alert('Error: ' + e.message);
}
```

### 3. Data Validation

```javascript
// Validate request data before processing
function validateRequest(request) {
  return request && 
         request.employeeId && 
         request.name && 
         request.start && 
         request.end;
}

data.forEach(function(request) {
  if (validateRequest(request)) {
    // Process valid request
  } else {
    console.log('Invalid request:', request);
  }
});
```

## Troubleshooting

### Common Issues

1. **"Received XML response instead of JSON"**
   - Check your API key and permissions
   - Verify your subdomain is correct

2. **"No time-off requests found"**
   - Check the date range
   - Verify API permissions include time-off access

3. **"An error occurred while fetching"**
   - Check network connectivity
   - Verify API key format
   - Review Google Apps Script quotas

### Debug Mode

```javascript
// Enable debug logging
function syncTimeOffRequestsDebug() {
  console.log('Starting sync...');
  console.log('API Key:', apiKey ? 'Set' : 'Not set');
  console.log('Subdomain:', subdomain);
  
  // ... rest of function
}
```

## Dependencies

- Google Apps Script runtime V8
- BambooHR API
- UrlFetchApp service

