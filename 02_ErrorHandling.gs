/**
 * YSL Hub v2 Error Handling Module
 * 
 * This module provides centralized error handling, logging, and user-friendly
 * error messages. It maintains a system log that can be viewed for debugging.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

// Log levels and their colors for visual indication
const LOG_LEVELS = {
  DEBUG: { value: 0, color: '#808080' },   // Gray
  INFO: { value: 1, color: '#000000' },    // Black
  WARNING: { value: 2, color: '#FFA500' }, // Orange
  ERROR: { value: 3, color: '#FF0000' }    // Red
};

// Maximum number of log entries to keep
const MAX_LOG_ENTRIES = 500;

// Current minimum log level to display (can be changed in settings)
let currentLogLevel = LOG_LEVELS.INFO.value;

/**
 * Initialize the error handling system
 * Creates a log sheet if it doesn't exist
 */
function initializeErrorHandling() {
  try {
    // Set log level based on script properties if available
    const storedLevel = PropertiesService.getScriptProperties().getProperty('logLevel');
    if (storedLevel && LOG_LEVELS[storedLevel]) {
      currentLogLevel = LOG_LEVELS[storedLevel].value;
    }
    
    // Create log sheet if it doesn't exist
    ensureLogSheetExists();
    
    // Log initialization success
    logMessage('Error handling system initialized', 'INFO');
    return true;
  } catch (error) {
    // Use native Logger as fallback
    Logger.log(`Error initializing error handling: ${error.message}`);
    return false;
  }
}

/**
 * Ensure the log sheet exists, create it if needed
 * @return {Sheet} The log sheet
 */
function ensureLogSheetExists() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName('SystemLog');
  
  // Create the log sheet if it doesn't exist
  if (!logSheet) {
    logSheet = ss.insertSheet('SystemLog');
    
    // Set up headers
    const headers = ['Timestamp', 'Level', 'Source', 'Message', 'Details'];
    logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    logSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Set column widths for better visibility
    logSheet.setColumnWidth(1, 180);  // Timestamp
    logSheet.setColumnWidth(2, 80);   // Level
    logSheet.setColumnWidth(3, 120);  // Source
    logSheet.setColumnWidth(4, 300);  // Message
    logSheet.setColumnWidth(5, 400);  // Details
    
    // Freeze header row
    logSheet.setFrozenRows(1);
    
    // Hide the sheet to avoid cluttering the interface
    logSheet.hideSheet();
  }
  
  return logSheet;
}

/**
 * Log a message to the system log
 * 
 * @param {string} message - The message to log
 * @param {string} level - Log level (DEBUG, INFO, WARNING, ERROR)
 * @param {string} source - Source of the message (function or module name)
 * @param {string} details - Additional details (optional)
 */
function logMessage(message, level = 'INFO', source = '', details = '') {
  // Default to INFO level if an invalid level is provided
  const logLevel = LOG_LEVELS[level] ? level : 'INFO';
  
  // Always log to native Logger first as backup
  Logger.log(`[${logLevel}] ${source ? `[${source}] ` : ''}${message} ${details ? `- ${details}` : ''}`);
  
  // Skip logging to sheet if level is below current threshold
  if (LOG_LEVELS[logLevel].value < currentLogLevel) {
    return;
  }
  
  try {
    // Get stack trace to determine source if not provided
    if (!source) {
      try {
        throw new Error();
      } catch (e) {
        const stackLines = e.stack.split('\n');
        // Look for the calling function in the stack trace
        if (stackLines.length >= 3) {
          // Format: "at functionName (fileName:line:column)"
          const callerLine = stackLines[2].trim();
          const functionMatch = callerLine.match(/at\s+([^\s(]+)/);
          if (functionMatch && functionMatch[1]) {
            source = functionMatch[1];
          }
        }
      }
    }
    
    // Get current timestamp
    const timestamp = new Date();
    
    // Add log entry to sheet
    const logSheet = ensureLogSheetExists();
    
    // Prepare log entry
    const logEntry = [
      Utilities.formatDate(timestamp, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
      logLevel,
      source,
      message,
      details
    ];
    
    // Insert at the top of the log (after header row)
    logSheet.insertRowAfter(1);
    logSheet.getRange(2, 1, 1, logEntry.length).setValues([logEntry]);
    
    // Set color based on log level
    logSheet.getRange(2, 2).setFontColor(LOG_LEVELS[logLevel].color);
    
    // Trim log if it exceeds maximum size
    const logRows = logSheet.getLastRow();
    if (logRows > MAX_LOG_ENTRIES + 1) { // +1 for header
      logSheet.deleteRows(MAX_LOG_ENTRIES + 2, logRows - (MAX_LOG_ENTRIES + 1));
    }
  } catch (error) {
    // Use native Logger as fallback
    Logger.log(`Error writing to log sheet: ${error.message}`);
  }
}

/**
 * Set the minimum log level for display
 * 
 * @param {string} level - Minimum log level (DEBUG, INFO, WARNING, ERROR)
 * @return {boolean} Success status
 */
function setLogLevel(level) {
  if (!LOG_LEVELS[level]) {
    logMessage(`Invalid log level: ${level}`, 'ERROR');
    return false;
  }
  
  currentLogLevel = LOG_LEVELS[level].value;
  PropertiesService.getScriptProperties().setProperty('logLevel', level);
  logMessage(`Log level set to ${level}`, 'INFO');
  return true;
}

/**
 * Handle an error with consistent formatting and logging
 * 
 * @param {Error|string} error - The error object or message
 * @param {string} source - Source of the error (function or module name)
 * @param {string} userMessage - User-friendly message (optional)
 * @param {boolean} showAlert - Whether to show an alert to the user
 * @return {Object} Error information object
 */
function handleError(error, source = '', userMessage = '', showAlert = true) {
  // Format error details
  const errorMessage = error instanceof Error ? error.message : error.toString();
  const errorDetails = error instanceof Error ? error.stack : '';
  
  // Log the error
  logMessage(errorMessage, 'ERROR', source, errorDetails);
  
  // Use a generic user message if none provided
  const displayMessage = userMessage || 'An error occurred while performing this operation.';
  
  // Show alert to user if requested
  if (showAlert) {
    try {
      const ui = SpreadsheetApp.getUi();
      ui.alert(
        'Operation Error',
        `${displayMessage}\n\nTechnical details: ${errorMessage}`,
        ui.ButtonSet.OK
      );
    } catch (alertError) {
      // Alert couldn't be shown, just log this additional error
      logMessage(`Could not show error alert: ${alertError.message}`, 'WARNING');
    }
  }
  
  // Return error information for potential further handling
  return {
    message: errorMessage,
    userMessage: displayMessage,
    source: source,
    timestamp: new Date().toISOString(),
    details: errorDetails
  };
}

/**
 * Display a user-friendly confirmation dialog
 * 
 * @param {string} title - Dialog title
 * @param {string} message - Dialog message
 * @param {boolean} includeCancel - Whether to include a Cancel button
 * @return {boolean} True if user confirmed, false otherwise
 */
function confirmAction(title, message, includeCancel = true) {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      title,
      message,
      includeCancel ? ui.ButtonSet.YES_NO : ui.ButtonSet.YES_NO_CANCEL
    );
    
    return result === ui.Button.YES;
  } catch (error) {
    logMessage(`Error showing confirmation dialog: ${error.message}`, 'ERROR');
    return false;
  }
}

/**
 * Show the log viewer to the user
 */
function showLogViewer() {
  try {
    // Get the log sheet and make it visible
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('SystemLog');
    
    if (!logSheet) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Log Not Found', 'The system log sheet was not found.', ui.ButtonSet.OK);
      return;
    }
    
    // Unhide the sheet
    logSheet.showSheet();
    
    // Activate the sheet
    logSheet.activate();
    
    // Select the first data row
    logSheet.setActiveRange(logSheet.getRange(2, 1));
    
    // Show instructions
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'System Log',
      'You are now viewing the system log. This shows all system events and errors.\n\n' +
      'To hide this sheet again, select "Hide Log" from the YSL Hub menu.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    logMessage(`Error showing log viewer: ${error.message}`, 'ERROR');
  }
}

/**
 * Hide the log sheet
 */
function hideLogSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('SystemLog');
    
    if (logSheet) {
      logSheet.hideSheet();
    }
  } catch (error) {
    // Just log the error without showing an alert, as this is a UI cleanup operation
    logMessage(`Error hiding log sheet: ${error.message}`, 'ERROR');
  }
}

/**
 * Clear the log sheet
 */
function clearLog() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Clear System Log',
      'Are you sure you want to clear all log entries? This cannot be undone.',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('SystemLog');
    
    if (!logSheet) {
      ui.alert('Log Not Found', 'The system log sheet was not found.', ui.ButtonSet.OK);
      return;
    }
    
    // Delete all rows except headers and add a fresh row
    const rowCount = logSheet.getLastRow();
    if (rowCount > 1) {
      logSheet.deleteRows(2, rowCount - 1);
    }
    
    // Add a log cleared entry
    logMessage('System log cleared', 'INFO');
    
    ui.alert('Log Cleared', 'The system log has been cleared.', ui.ButtonSet.OK);
  } catch (error) {
    logMessage(`Error clearing log: ${error.message}`, 'ERROR');
  }
}

/**
 * Export the log to a new spreadsheet for archiving
 * @return {string} URL of the exported spreadsheet or empty string on failure
 */
function exportLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName('SystemLog');
    
    if (!logSheet) {
      const ui = SpreadsheetApp.getUi();
      ui.alert('Log Not Found', 'The system log sheet was not found.', ui.ButtonSet.OK);
      return '';
    }
    
    // Create a new spreadsheet for the log
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const exportSS = SpreadsheetApp.create(`YSL Hub System Log Export ${timestamp}`);
    const exportSheet = exportSS.getActiveSheet();
    
    // Copy data from log sheet to export sheet
    const logData = logSheet.getDataRange().getValues();
    exportSheet.getRange(1, 1, logData.length, logData[0].length).setValues(logData);
    
    // Format the export sheet
    exportSheet.getRange(1, 1, 1, logData[0].length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Set column widths for better visibility
    exportSheet.setColumnWidth(1, 180);  // Timestamp
    exportSheet.setColumnWidth(2, 80);   // Level
    exportSheet.setColumnWidth(3, 120);  // Source
    exportSheet.setColumnWidth(4, 300);  // Message
    exportSheet.setColumnWidth(5, 400);  // Details
    
    // Freeze header row
    exportSheet.setFrozenRows(1);
    
    // Get the URL of the exported spreadsheet
    const exportUrl = exportSS.getUrl();
    
    // Show confirmation with the URL
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Log Exported',
      `The system log has been exported to a new spreadsheet.\n\nURL: ${exportUrl}`,
      ui.ButtonSet.OK
    );
    
    // Log the export
    logMessage(`System log exported to ${exportUrl}`, 'INFO');
    
    return exportUrl;
  } catch (error) {
    logMessage(`Error exporting log: ${error.message}`, 'ERROR');
    return '';
  }
}

// Make functions available to other modules
const ErrorHandling = {
  initializeErrorHandling: initializeErrorHandling,
  logMessage: logMessage,
  handleError: handleError,
  confirmAction: confirmAction,
  showLogViewer: showLogViewer,
  hideLogSheet: hideLogSheet,
  clearLog: clearLog,
  exportLog: exportLog,
  setLogLevel: setLogLevel
};