/**
 * BambooHR Time Off Request Sync (Enhanced Version)
 * 
 * This script syncs time-off requests from BambooHR to a dedicated sheet.
 * Replace the placeholder values with your actual BambooHR credentials.
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('BambooHR')
    .addItem('Sync Time Off Requests', 'syncTimeOffRequests')
    .addToUi();
}

/**
 * Calculates working days between two dates
 * @param {string} startDate - Start date
 * @param {string} endDate - End date
 * @returns {number} Total calculated days
 */
function calculateDays(startDate, endDate) {
  var start = new Date(startDate);
  var end = new Date(endDate);
  var totalDays = 0;
  var currentDate = new Date(start);
  
  while (currentDate <= end) {
    var dayOfWeek = currentDate.getDay();
    // 0 is Sunday, 6 is Saturday
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      totalDays += 1.69; // Weekend day
    } else {
      totalDays += 1.25; // Weekday
    }
    currentDate.setDate(currentDate.getDate() + 1);
  }
  
  return totalDays;
}

/**
 * Syncs time-off requests from BambooHR to a dedicated sheet
 */
function syncTimeOffRequests() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName('Time Off Requests');
  
  // Create the sheet if it doesn't exist
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Time Off Requests');
  }
  
  sheet.clear();

  // Configuration - Replace with your actual values
  var apiKey = 'YOUR_BAMBOOHR_API_KEY_HERE';
  var subdomain = 'YOUR_BAMBOOHR_SUBDOMAIN_HERE';
  
  var today = new Date();
  var oneMonthAgo = new Date();
  oneMonthAgo.setMonth(today.getMonth() - 1);

  var startDate = oneMonthAgo.toISOString().split('T')[0];
  var endDate = today.toISOString().split('T')[0];

  var url = 'https://' + subdomain + '.bamboohr.com/api/gateway.php/' + subdomain + '/v1/time_off/requests?start=' + startDate + '&end=' + endDate;

  var options = {
    method: 'get',
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(apiKey + ':x'),
      Accept: "application/json"
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var content = response.getContentText();

    if (content.trim().startsWith('<?xml')) {
      SpreadsheetApp.getUi().alert('Error: Received an unexpected response from BambooHR. Check the logs.');
      return;
    }

    var data = JSON.parse(content);

    if (!data || data.length === 0) {
      sheet.appendRow(['No time-off requests found']);
      return;
    }

    var headers = ['Request Created Date', 'Employee ID', 'Name', 'Start Date', 'End Date', 'Status', 'Type', 'Reason', 'Amount', 'Unit'];
    sheet.appendRow(headers);

    // Format headers with white text and dark blue background
    var headerRange = sheet.getRange(sheet.getLastRow(), 1, 1, headers.length);
    headerRange.setFontColor('white');
    headerRange.setBackground('#1f4e79'); // Dark blue color
    headerRange.setFontWeight('bold');

    var processedCount = 0;
    data.forEach(function(request, index) {
      try {
        // Extract amount and unit information
        var amount = '';
        var unit = '';
        if (request.amount) {
          amount = request.amount.amount;
          unit = request.amount.unit;
        }

        // Extract request creation date (when request was submitted)
        var requestCreatedDate = '';
        if (request.created) {
          var createdDate = new Date(request.created);
          requestCreatedDate = createdDate.toLocaleDateString();
        }

        sheet.appendRow([
          requestCreatedDate,
          request.employeeId,
          request.name,
          request.start,
          request.end,
          request.status.status,
          request.type.name,
          request.notes.employee || '',
          amount,
          unit
        ]);
        
        processedCount++;
        
      } catch (e) {
        // Silent error handling for individual requests
      }
    });

    SpreadsheetApp.getUi().alert('Time-off requests synced successfully!');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert('An error occurred while fetching time-off requests. Check logs for details.');
  }
} 