// =====================================================
// 17_UsageTracking.js — User Activity Tracking
// Logs user actions to track who's using the system
// =====================================================

const USAGE_LOG_SHEET_NAME = 'Usage_Log';

/**
 * Creates the Usage_Log sheet if it doesn't exist.
 */
function setupUsageLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let usageSheet = ss.getSheetByName(USAGE_LOG_SHEET_NAME);

  if (usageSheet) {
    return `Usage_Log sheet already exists.`;
  }

  // Create the sheet
  usageSheet = ss.insertSheet(USAGE_LOG_SHEET_NAME);

  // Set up headers
  const headers = [
    'Timestamp',
    'User Email',
    'User Name',
    'Action',
    'Details',
    'Status',
    'Duration (sec)'
  ];

  usageSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  usageSheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');
  usageSheet.setFrozenRows(1);

  // Auto-resize columns
  usageSheet.autoResizeColumns(1, headers.length);

  // Set column widths for better readability
  usageSheet.setColumnWidth(1, 180); // Timestamp
  usageSheet.setColumnWidth(2, 200); // Email
  usageSheet.setColumnWidth(3, 150); // Name
  usageSheet.setColumnWidth(4, 200); // Action
  usageSheet.setColumnWidth(5, 300); // Details
  usageSheet.setColumnWidth(6, 100); // Status
  usageSheet.setColumnWidth(7, 100); // Duration

  const msg = `✅ Created Usage_Log sheet for user activity tracking.`;
  ss.toast(msg, 'Usage Tracking', 6);
  logImportEvent('Usage Log', 'Created Usage_Log sheet');

  return msg;
}

/**
 * Logs a user action to the Usage_Log sheet.
 * Automatically captures user email, name, and timestamp.
 */
function logUserAction(action, details, status, durationSeconds) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let usageSheet = ss.getSheetByName(USAGE_LOG_SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!usageSheet) {
      setupUsageLogSheet();
      usageSheet = ss.getSheetByName(USAGE_LOG_SHEET_NAME);
    }

    // Get user info
    const userEmail = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || 'Unknown';
    const userName = userEmail.split('@')[0] || 'Unknown';

    // Prepare row data
    const timestamp = new Date();
    const statusValue = status || 'Success';
    const duration = durationSeconds !== undefined ? durationSeconds : '';

    const rowData = [
      timestamp,
      userEmail,
      userName,
      action || 'Unknown Action',
      details || '',
      statusValue,
      duration
    ];

    // Append to sheet
    usageSheet.appendRow(rowData);

    // Format the new row
    const lastRow = usageSheet.getLastRow();
    usageSheet.getRange(lastRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    // Color code by status
    if (statusValue === 'Error' || statusValue === 'Failed') {
      usageSheet.getRange(lastRow, 6).setBackground('#fce8e6'); // Light red
    } else if (statusValue === 'Success') {
      usageSheet.getRange(lastRow, 6).setBackground('#d9ead3'); // Light green
    }

  } catch (e) {
    // If logging fails, don't break the main function
    console.error('Failed to log user action: ' + e.message);
  }
}

/**
 * Wrapper function to track execution of any function.
 * Usage: trackUserAction('Import Orders', () => importShopifyOrders())
 */
function trackUserAction(actionName, functionToRun, details) {
  const startTime = new Date();
  let status = 'Success';
  let result = null;

  try {
    result = functionToRun();
    logUserAction(actionName, details, 'Success', (new Date() - startTime) / 1000);
    return result;
  } catch (e) {
    status = 'Error';
    logUserAction(actionName, details || e.message, 'Error', (new Date() - startTime) / 1000);
    throw e;
  }
}

/**
 * Enhanced versions of main functions with usage tracking.
 * These can be called from the sidebar or menu.
 */
function importShopifyOrdersTracked() {
  return trackUserAction('Import Shopify Orders', () => importShopifyOrders(), 'Last 14 days');
}

function importSquarespaceOrdersTracked() {
  return trackUserAction('Import Squarespace Orders', () => importSquarespaceOrders(), 'Last 14 days');
}

function updateShopifyOrdersWithRefundsTracked() {
  return trackUserAction('Update Shopify Refunds', () => updateShopifyOrdersWithRefunds(), 'Last 90 days');
}

function updateSquarespaceOrdersWithRefundsTracked() {
  return trackUserAction('Update Squarespace Refunds', () => updateSquarespaceOrdersWithRefunds(), 'Last 90 days');
}

function buildAllOrdersCleanTracked() {
  return trackUserAction('Build All Orders Clean', () => buildAllOrdersClean());
}

function buildOrdersSummaryReportTracked() {
  return trackUserAction('Build Summary Report', () => buildOrdersSummaryReport());
}

function buildRefundsReportTracked() {
  return trackUserAction('Build Refunds Report', () => buildRefundsReport());
}

function buildDiscountsReportTracked() {
  return trackUserAction('Build Discounts Report', () => buildDiscountsReport());
}

function buildCustomerOutreachListTracked() {
  return trackUserAction('Build Outreach List', () => buildCustomerOutreachList());
}

function importAndUpdateAllOrdersTracked() {
  return trackUserAction('Import & Update All Orders', () => importAndUpdateAllOrders(), 'Full workflow');
}

function addToBannedListTracked(emailOrDomain) {
  return trackUserAction('Add to Banned List', () => addToBannedList(emailOrDomain), `Banned: ${emailOrDomain}`);
}

/**
 * Gets usage statistics for display.
 */
function getUsageStats() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(USAGE_LOG_SHEET_NAME);

  if (!usageSheet || usageSheet.getLastRow() < 2) {
    return {
      totalActions: 0,
      uniqueUsers: 0,
      lastAction: 'Never',
      mostActiveUser: 'N/A'
    };
  }

  const data = usageSheet.getRange(2, 1, usageSheet.getLastRow() - 1, 7).getValues();

  const uniqueUsers = new Set();
  const userCounts = {};
  let lastTimestamp = null;

  data.forEach(row => {
    const email = row[1];
    if (email) {
      uniqueUsers.add(email);
      userCounts[email] = (userCounts[email] || 0) + 1;
    }

    const timestamp = row[0];
    if (timestamp && (!lastTimestamp || timestamp > lastTimestamp)) {
      lastTimestamp = timestamp;
    }
  });

  // Find most active user
  let mostActiveUser = 'N/A';
  let maxCount = 0;
  for (const [user, count] of Object.entries(userCounts)) {
    if (count > maxCount) {
      maxCount = count;
      mostActiveUser = user;
    }
  }

  return {
    totalActions: data.length,
    uniqueUsers: uniqueUsers.size,
    lastAction: lastTimestamp ? Utilities.formatDate(lastTimestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : 'Never',
    mostActiveUser: mostActiveUser + ' (' + maxCount + ' actions)'
  };
}

/**
 * Clears old usage logs (older than specified days).
 */
function clearOldUsageLogs(daysToKeep) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const usageSheet = ss.getSheetByName(USAGE_LOG_SHEET_NAME);

  if (!usageSheet || usageSheet.getLastRow() < 2) {
    return 'No usage logs to clear.';
  }

  daysToKeep = daysToKeep || 90; // Default 90 days
  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);

  const data = usageSheet.getRange(2, 1, usageSheet.getLastRow() - 1, 7).getValues();
  const rowsToKeep = [usageSheet.getRange(1, 1, 1, 7).getValues()[0]]; // Keep header

  let deletedCount = 0;

  data.forEach(row => {
    const timestamp = row[0];
    if (timestamp && timestamp >= cutoffDate) {
      rowsToKeep.push(row);
    } else {
      deletedCount++;
    }
  });

  if (deletedCount > 0) {
    usageSheet.clearContents();
    usageSheet.getRange(1, 1, rowsToKeep.length, 7).setValues(rowsToKeep);

    // Reapply formatting
    usageSheet.getRange(1, 1, 1, 7)
      .setFontWeight('bold')
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
    usageSheet.setFrozenRows(1);
  }

  const msg = `Cleared ${deletedCount} usage log entries older than ${daysToKeep} days.`;
  ss.toast(msg, 'Usage Log Cleanup', 6);
  return msg;
}
