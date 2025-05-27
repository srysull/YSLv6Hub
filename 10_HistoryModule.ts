/**
 * YSL Hub v2 History Module
 * 
 * This module provides functionality for tracking and displaying the history
 * of system changes, updates, and important events in the YSL Hub system.
 * It maintains a chronological record of activities to facilitate system
 * monitoring and accountability.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

// History sheet name
const HISTORY_SHEET_NAME = 'SystemHistory';

// Event types
const EVENT_TYPES = {
  INITIALIZATION: 'System Initialization',
  CONFIGURATION: 'Configuration Change',
  DATA_IMPORT: 'Data Import',
  REPORT_GENERATION: 'Report Generation',
  COMMUNICATION: 'Communication',
  SESSION_TRANSITION: 'Session Transition',
  SYSTEM_UPDATE: 'System Update',
  USER_ACTION: 'User Action',
  VERSION_UPDATE: 'Version Update'
};

/**
 * Creates and displays a history sheet
 * 
 * @returns Success status
 */
function createHistorySheet() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Creating history sheet', 'INFO', 'createHistorySheet');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    
    // Create the sheet if it doesn't exist
    if (!historySheet) {
      historySheet = ss.insertSheet(HISTORY_SHEET_NAME);
      
      // Set up headers
      const headers = [
        'Date', 'Time', 'Event Type', 'User', 'Description', 'Details', 'Version'
      ];
      
      historySheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      historySheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      // Set column widths for better visibility
      historySheet.setColumnWidth(1, 100);  // Date
      historySheet.setColumnWidth(2, 80);   // Time
      historySheet.setColumnWidth(3, 150);  // Event Type
      historySheet.setColumnWidth(4, 150);  // User
      historySheet.setColumnWidth(5, 250);  // Description
      historySheet.setColumnWidth(6, 300);  // Details
      historySheet.setColumnWidth(7, 100);  // Version
      
      // Freeze header row
      historySheet.setFrozenRows(1);
      
      // Create data validation for event type
      const typeRange = historySheet.getRange(2, 3, 100, 1);
      const typeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(Object.values(EVENT_TYPES), true)
        .build();
      typeRange.setDataValidation(typeRule);
      
      // Add initial system history
      loadInitialHistory(historySheet);
    }
    
    // Ensure the history sheet is visible and active
    historySheet.showSheet();
    historySheet.activate();
    
    SpreadsheetApp.getUi().alert(
      'History View',
      'The System History sheet displays a chronological record of system events and changes.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createHistorySheet', 
        'Error creating history sheet. Please try again or contact support.');
    } else {
      Logger.log(`Error creating history sheet: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create history sheet: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Loads initial history data
 * 
 * @param historySheet - The history sheet
 */
function loadInitialHistory(historySheet) {
  try {
    // Get version information
    let currentVersion = '2.0.0';
    let versionDate = '2025-04-27';
    
    if (VersionControl && typeof VersionControl.getVersionInfo === 'function') {
      const versionInfo = VersionControl.getVersionInfo();
      currentVersion = versionInfo.currentVersion;
      versionDate = versionInfo.releaseDate;
    }
    
    // Check for existing version update records in script properties
    const scriptProps = PropertiesService.getScriptProperties();
    const versionHistory = scriptProps.getProperty('versionHistory');
    
    let historyEntries = [];
    
    if (versionHistory) {
      try {
        historyEntries = JSON.parse(versionHistory);
      } catch (e) {
        historyEntries = [];
      }
    }
    
    // If no history, add system initialization record
    if (historyEntries.length === 0) {
      historyEntries.push({
        date: versionDate,
        time: '08:00:00',
        type: EVENT_TYPES.INITIALIZATION,
        user: 'System',
        description: 'YSL Hub system installed',
        details: `Initial system installation with version ${currentVersion}`,
        version: currentVersion
      });
    }
    
    // Add any version update records
    const versionUpdates = scriptProps.getProperty('versionUpdates');
    if (versionUpdates) {
      try {
        const updates = JSON.parse(versionUpdates);
        for (const update of updates) {
          historyEntries.push({
            date: update.date,
            time: update.time || '12:00:00',
            type: EVENT_TYPES.VERSION_UPDATE,
            user: 'System',
            description: `Updated to version ${update.version}`,
            details: update.details || 'System version update',
            version: update.version
          });
        }
      } catch (e) {
        // Ignore parsing errors
      }
    }
    
    // Add any configuration changes
    const configChanges = scriptProps.getProperty('configChanges');
    if (configChanges) {
      try {
        const changes = JSON.parse(configChanges);
        for (const change of changes) {
          historyEntries.push({
            date: change.date,
            time: change.time || '12:00:00',
            type: EVENT_TYPES.CONFIGURATION,
            user: change.user || 'Administrator',
            description: 'Configuration updated',
            details: change.details || 'System configuration changed',
            version: change.version || currentVersion
          });
        }
      } catch (e) {
        // Ignore parsing errors
      }
    }
    
    // Get session history
    const sessionHistory = scriptProps.getProperty('sessionHistory');
    if (sessionHistory) {
      try {
        const sessions = JSON.parse(sessionHistory);
        for (const session of sessions) {
          historyEntries.push({
            date: session.date,
            time: session.time || '12:00:00',
            type: EVENT_TYPES.SESSION_TRANSITION,
            user: session.user || 'Administrator',
            description: `Transitioned to session: ${session.name}`,
            details: session.details || 'New session started',
            version: session.version || currentVersion
          });
        }
      } catch (e) {
        // Ignore parsing errors
      }
    }
    
    // Add user action records from logs
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = ss.getSheetByName('SystemLog');
      
      if (logSheet) {
        const logData = logSheet.getDataRange().getValues();
        
        // Skip header row
        for (let i = 1; i < Math.min(50, logData.length); i++) {
          const row = logData[i];
          
          // Only add significant log entries (INFO level or higher)
          if (row[1] === 'INFO' || row[1] === 'WARNING' || row[1] === 'ERROR') {
            // Parse timestamp into date and time
            let dateStr = '';
            let timeStr = '';
            
            if (row[0]) {
              const timestamp = row[0].toString();
              const parts = timestamp.split(' ');
              
              if (parts.length >= 2) {
                dateStr = parts[0];
                timeStr = parts[1];
              } else {
                dateStr = timestamp;
                timeStr = '12:00:00';
              }
            } else {
              const now = new Date();
              dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
              timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
            }
            
            historyEntries.push({
              date: dateStr,
              time: timeStr,
              type: determineEventType(row[2], row[3]),
              user: 'System',
              description: row[3] || 'System Event',
              details: row[4] || '',
              version: currentVersion
            });
          }
        }
      }
    } catch (e) {
      // Ignore errors reading log data
    }
    
    // Sort entries by date and time
    historyEntries.sort((a, b) => {
      const dateA = a.date + ' ' + a.time;
      const dateB = b.date + ' ' + b.time;
      return dateA.localeCompare(dateB);
    });
    
    // Add to history sheet
    if (historyEntries.length > 0) {
      const data = historyEntries.map(entry => [
        entry.date,
        entry.time,
        entry.type,
        entry.user,
        entry.description,
        entry.details,
        entry.version
      ]);
      
      historySheet.getRange(2, 1, data.length, 7).setValues(data);
    }
  } catch (error) {
    Logger.log(`Error loading initial history: ${error.message}`);
  }
}

/**
 * Determines the event type based on source and message
 * 
 * @param source - The source of the event
 * @param message - The event message
 * @returns The determined event type
 */
function determineEventType(source, message) {
  if (!source && !message) {
    return EVENT_TYPES.USER_ACTION;
  }
  
  const sourceStr = (source || '').toLowerCase();
  const messageStr = (message || '').toLowerCase();
  
  if (sourceStr.includes('version') || messageStr.includes('version') || 
      messageStr.includes('updated to')) {
    return EVENT_TYPES.VERSION_UPDATE;
  }
  
  if (sourceStr.includes('init') || messageStr.includes('initial') || 
      messageStr.includes('initialized')) {
    return EVENT_TYPES.INITIALIZATION;
  }
  
  if (sourceStr.includes('config') || messageStr.includes('configuration') || 
      messageStr.includes('setting')) {
    return EVENT_TYPES.CONFIGURATION;
  }
  
  if (sourceStr.includes('import') || messageStr.includes('import') || 
      messageStr.includes('data') || messageStr.includes('sync')) {
    return EVENT_TYPES.DATA_IMPORT;
  }
  
  if (sourceStr.includes('report') || messageStr.includes('report') || 
      messageStr.includes('generated')) {
    return EVENT_TYPES.REPORT_GENERATION;
  }
  
  if (sourceStr.includes('email') || sourceStr.includes('communicate') || 
      messageStr.includes('email') || messageStr.includes('sent')) {
    return EVENT_TYPES.COMMUNICATION;
  }
  
  if (sourceStr.includes('session') || messageStr.includes('session')) {
    return EVENT_TYPES.SESSION_TRANSITION;
  }
  
  if (sourceStr.includes('system') || messageStr.includes('system')) {
    return EVENT_TYPES.SYSTEM_UPDATE;
  }
  
  return EVENT_TYPES.USER_ACTION;
}

/**
 * Records a history event
 * 
 * @param eventType - The type of event
 * @param description - Description of the event
 * @param details - Additional details
 * @param user - The user who initiated the event
 * @returns Success status
 */
function recordHistoryEvent(eventType, description, details = '', user = '') {
  try {
    // Get the current version
    let currentVersion = '2.0.0';
    
    if (VersionControl && typeof VersionControl.getCurrentVersion === 'function') {
      currentVersion = VersionControl.getCurrentVersion();
    }
    
    // Get the current user if not provided
    if (!user) {
      user = Session.getEffectiveUser().getEmail();
    }
    
    // Current timestamp
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    // Create history entry
    const historyEntry = [
      dateStr,
      timeStr,
      eventType,
      user,
      description,
      details,
      currentVersion
    ];
    
    // Find the history sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    
    // If history sheet exists, add the entry
    if (historySheet) {
      historySheet.insertRowAfter(1);
      historySheet.getRange(2, 1, 1, historyEntry.length).setValues([historyEntry]);
    }
    
    // Log the event regardless
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`History Event: ${description}`, 'INFO', 'recordHistoryEvent', details);
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error recording history event: ${error.message}`);
    return false;
  }
}

/**
 * Exports the history to a Google Doc for archiving
 * 
 * @returns The URL of the exported document or empty string on failure
 */
function exportHistory() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Exporting history', 'INFO', 'exportHistory');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    
    if (!historySheet) {
      SpreadsheetApp.getUi().alert(
        'History Not Found',
        'The System History sheet was not found. Please create it first using "View History".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return '';
    }
    
    // Get system and version info for the document header
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || 'Current Session';
    
    let versionInfo = 'Version 2.0.0 (2025-04-27)';
    if (VersionControl && typeof VersionControl.getVersionInfo === 'function') {
      const vi = VersionControl.getVersionInfo();
      versionInfo = `Version ${vi.currentVersion} (${vi.releaseDate})`;
    }
    
    // Create a new Google Doc for the history
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const docName = `YSL Hub System History Export - ${timestamp}`;
    const doc = DocumentApp.create(docName);
    
    // Get the document body
    const body = doc.getBody();
    
    // Add title
    body.appendParagraph('YSL Hub System History')
      .setHeading(DocumentApp.ParagraphHeading.HEADING1)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Add system info
    body.appendParagraph(`Session: ${sessionName}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    body.appendParagraph(`${versionInfo}`)
      .setHeading(DocumentApp.ParagraphHeading.HEADING2)
      .setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    const exportParagraph = body.appendParagraph(`Exported on: ${timestamp.replace('_', ' ')}`);
    exportParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    exportParagraph.editAsText().setItalic(true);
    
    // Add spacing
    body.appendParagraph('');
    
    // Get history data
    const historyData = historySheet.getDataRange().getValues();
    
    // Create a table for the history
    const numRows = historyData.length;
    const numCols = historyData[0].length;
    
    // Create table
    const table = body.appendTable(historyData);
    
    // Format header row
    const headerRow = table.getRow(0);
    for (let i = 0; i < numCols; i++) {
      headerRow.getCell(i).setBackgroundColor('#f3f3f3');
      headerRow.getCell(i).editAsText().setBold(true);
    }
    
    // Save the document
    doc.saveAndClose();
    
    // Log the export
    recordHistoryEvent(
      EVENT_TYPES.USER_ACTION,
      'History Exported',
      `System history exported to document: ${docName}`
    );
    
    // Show confirmation with the URL
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'History Exported',
      `The system history has been exported to a new document.\n\nURL: ${doc.getUrl()}`,
      ui.ButtonSet.OK
    );
    
    return doc.getUrl();
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'exportHistory', 
        'Error exporting history. Please try again or contact support.');
    } else {
      Logger.log(`Error exporting history: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Export Failed',
        `Failed to export history: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return '';
  }
}

/**
 * Filters the history sheet to show specific types of events
 * 
 * @param eventType - The type of event to filter for, or 'All' for all events
 * @returns Success status
 */
function filterHistory(eventType) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Filtering history by type: ${eventType}`, 'INFO', 'filterHistory');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    
    if (!historySheet) {
      SpreadsheetApp.getUi().alert(
        'History Not Found',
        'The System History sheet was not found. Please create it first using "View History".',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Remove any existing filter
    historySheet.getFilter()?.remove();
    
    // If showing all events, just remove filter and return
    if (eventType === 'All') {
      return true;
    }
    
    // Create a filter on the event type column (column 3)
    const range = historySheet.getDataRange();
    const filter = range.createFilter();
    
    // Set filter criteria
    const typeColumn = 3;
    const criteria = SpreadsheetApp.newFilterCriteria()
      .whenTextEqualTo(eventType)
      .build();
    
    filter.setColumnFilterCriteria(typeColumn, criteria);
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error filtering history: ${error.message}`, 'ERROR', 'filterHistory');
    } else {
      Logger.log(`Error filtering history: ${error.message}`);
    }
    return false;
  }
}

/**
 * Clears the history sheet after confirmation
 * 
 * @returns Success status
 */
function clearHistory() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Clearing history', 'INFO', 'clearHistory');
    }
    
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Clear History',
      'Are you sure you want to clear the entire history? This cannot be undone.',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return false;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const historySheet = ss.getSheetByName(HISTORY_SHEET_NAME);
    
    if (!historySheet) {
      ui.alert(
        'History Not Found',
        'The System History sheet was not found.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Export the history before clearing
    exportHistory();
    
    // Remove any existing filter
    historySheet.getFilter()?.remove();
    
    // Clear all rows except the header
    const lastRow = historySheet.getLastRow();
    if (lastRow > 1) {
      historySheet.deleteRows(2, lastRow - 1);
    }
    
    // Record the clearing action
    recordHistoryEvent(
      EVENT_TYPES.USER_ACTION,
      'History Cleared',
      'System history was cleared after export'
    );
    
    ui.alert(
      'History Cleared',
      'The system history has been cleared after exporting to a document.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'clearHistory', 
        'Error clearing history. Please try again or contact support.');
    } else {
      Logger.log(`Error clearing history: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Clear Failed',
        `Failed to clear history: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

// Global variable export
const HistoryModule = {
  createHistorySheet,
  recordHistoryEvent,
  exportHistory,
  filterHistory,
  clearHistory,
  EVENT_TYPES
};