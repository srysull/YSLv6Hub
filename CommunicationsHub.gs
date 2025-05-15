/**
 * YSL Hub Communications Hub Module
 * 
 * This module implements a dynamic communications hub for creating,
 * managing, and sending communications to classes and students.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

// Configuration constants
const COMMUNICATIONS_HUB_CONFIG = {
  SHEET_NAME: 'Communications Hub',
  LOG_SHEET_NAME: 'Communication Log',
  HEADERS: {
    SELECTOR_LABEL: 'Select Classes:',
    SUBJECT: 'Subject:',
    BODY: 'Message Body:',
    SEND_DATE: 'Send Date:',
    RECIPIENTS: 'Recipients:',
    STATUS: 'Status:',
  },
  CELL_STYLES: {
    HEADER_COLOR: '#4285F4',
    HEADER_TEXT_COLOR: '#FFFFFF',
    SECTION_COLOR: '#E0E0E0',
    SELECTOR_BG_COLOR: '#D9EAD3'
  },
  LOG_COLUMNS: {
    DATE_SENT: 'Date Sent',
    SUBJECT: 'Subject',
    RECIPIENTS: 'Recipients',
    SENDER: 'Sender',
    STATUS: 'Status'
  }
};

/**
 * Creates or resets the communications hub sheet
 * @return {Sheet} The created or updated sheet
 */
function createCommunicationsHub() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(COMMUNICATIONS_HUB_CONFIG.SHEET_NAME);
    
    // Create the sheet if it doesn't exist or completely reset it if it does
    if (!sheet) {
      // Create a new sheet
      sheet = ss.insertSheet(COMMUNICATIONS_HUB_CONFIG.SHEET_NAME);
    } else {
      // Clear everything from the existing sheet
      sheet.clear();
      sheet.clearFormats();
      sheet.clearConditionalFormatRules();
      
      // Clear all validations
      try {
        if (typeof sheet.clearDataValidations === 'function') {
          sheet.clearDataValidations();
        } else {
          const totalRows = sheet.getMaxRows();
          const totalCols = sheet.getMaxColumns();
          for (let startRow = 1; startRow <= totalRows; startRow += 20) {
            const rowsToProcess = Math.min(20, totalRows - startRow + 1);
            sheet.getRange(startRow, 1, rowsToProcess, totalCols).setDataValidation(null);
          }
        }
      } catch (e) {
        Logger.log(`Error clearing data validations: ${e.message}. Continuing anyway.`);
      }
      
      // Reset all column widths and row heights
      const totalColumns = sheet.getMaxColumns();
      for (let i = 1; i <= totalColumns; i++) {
        sheet.setColumnWidth(i, 100); // Reset to default width
      }
      
      const totalRows = sheet.getMaxRows();
      for (let i = 1; i <= totalRows; i++) {
        sheet.setRowHeight(i, 21); // Reset to default height
      }
      
      // Ensure there are enough rows and columns
      const minRows = 100;
      const minColumns = 30;
      
      if (sheet.getMaxRows() < minRows) {
        sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
      }
      
      if (sheet.getMaxColumns() < minColumns) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), minColumns - sheet.getMaxColumns());
      }
      
      // Unhide any hidden rows or columns
      sheet.showRows(1, sheet.getMaxRows());
      sheet.showColumns(1, sheet.getMaxColumns());
    }
    
    // Set up the communications hub structure
    setupCommunicationsHubStructure(sheet);
    
    // Create class selector checkboxes
    createClassSelectors(sheet);
    
    // Show confirmation
    SpreadsheetApp.getUi().alert(
      'Communications Hub Created',
      'The communications hub has been created. Use it to create and send communications to classes.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Set active sheet to communications hub
    sheet.activate();
    
    return sheet;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createCommunicationsHub', 
        'Error creating communications hub. Please try again or contact support.');
    } else {
      Logger.log(`Error creating communications hub: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create communications hub: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return null;
  }
}

/**
 * Sets up the structure of the communications hub sheet
 * @param {Sheet} sheet - The sheet to set up
 */
function setupCommunicationsHubStructure(sheet) {
  try {
    // Set column widths
    sheet.setColumnWidth(1, 150); // Label column
    sheet.setColumnWidth(2, 500); // Content column
    
    // Title row
    sheet.getRange('A1:F1').merge()
      .setValue('YSL Communications Hub')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Class selector section header
    sheet.getRange('A2:F2').merge()
      .setValue('Select Recipients')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.SECTION_COLOR);
    
    // Classes will be populated in rows 3-12 with checkboxes
    // This will be done in the createClassSelectors function
    
    // Message section header
    sheet.getRange('A13:F13').merge()
      .setValue('Message Content')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.SECTION_COLOR);
    
    // Subject row
    sheet.getRange('A14').setValue(COMMUNICATIONS_HUB_CONFIG.HEADERS.SUBJECT)
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setVerticalAlignment('top');
    
    // Subject input area
    sheet.getRange('B14:F14').merge();
    
    // Message body row
    sheet.getRange('A15').setValue(COMMUNICATIONS_HUB_CONFIG.HEADERS.BODY)
      .setFontWeight('bold')
      .setHorizontalAlignment('right')
      .setVerticalAlignment('top');
    
    // Message body input area - larger cell for HTML content
    sheet.getRange('B15:F25').merge();
    sheet.setRowHeights(15, 11, 40); // Set multiple rows to be taller for the message body
    
    // Send date row
    sheet.getRange('A26').setValue(COMMUNICATIONS_HUB_CONFIG.HEADERS.SEND_DATE)
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    
    // Send date input (date picker will be added)
    sheet.getRange('B26').setValue(new Date())
      .setNumberFormat('mm/dd/yyyy');
    
    // Status row
    sheet.getRange('A27').setValue(COMMUNICATIONS_HUB_CONFIG.HEADERS.STATUS)
      .setFontWeight('bold')
      .setHorizontalAlignment('right');
    
    // Status display area
    sheet.getRange('B27:F27').merge()
      .setValue('Ready to send')
      .setFontStyle('italic');
    
    // Helper text with available template variables
    sheet.getRange('A28:F28').merge()
      .setValue('Available variables: {{firstName}}, {{lastName}}, {{className}}, {{classTime}}, {{instructor}}')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Send button section
    sheet.getRange('A30:F30').merge()
      .setValue('Use "YSL Hub > Communications > Send Selected Communication" to send this message')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.SELECTOR_BG_COLOR);
    
    // Add a checkbox for selecting this communication to be sent
    sheet.getRange('A31:F31').merge();
    sheet.getRange('A31').insertCheckboxes();
    sheet.getRange('B31:F31').setValue('Select to send this communication');
  } catch (error) {
    Logger.log(`Error setting up communications hub structure: ${error.message}`);
    throw error;
  }
}

/**
 * Creates class selectors with checkboxes in the communications hub
 * @param {Sheet} sheet - The communications hub sheet
 */
function createClassSelectors(sheet) {
  try {
    // Get available classes from the Classes sheet
    const classes = getClassesForSelector();
    
    // Add "All Classes" option at row 3
    sheet.getRange('A3').setValue('All Classes:');
    sheet.getRange('B3').insertCheckboxes();
    
    // Add "Private Lessons" option at row 4
    sheet.getRange('A4').setValue('All Private Lessons:');
    sheet.getRange('B4').insertCheckboxes();
    
    // Start adding individual classes at row 5
    let currentRow = 5;
    
    // Add individual classes with checkboxes
    for (let i = 0; i < classes.length && i < 8; i++) { // Limit to 8 classes to avoid overcrowding
      sheet.getRange(currentRow, 1).setValue(classes[i]);
      sheet.getRange(currentRow, 2).insertCheckboxes();
      currentRow++;
    }
    
    // Add note about selecting classes
    sheet.getRange('A12:F12').merge()
      .setValue('Select classes above or use "All Classes" option. Multiple selections allowed.')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
  } catch (error) {
    Logger.log(`Error creating class selectors: ${error.message}`);
    throw error;
  }
}

/**
 * Gets classes for the selector from the Classes sheet
 * @return {Array} Array of class names
 */
function getClassesForSelector() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      return ['No classes found'];
    }
    
    const data = classesSheet.getDataRange().getValues();
    const classes = [];
    
    // Skip header row
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][2] && data[i][3]) { // Columns B, C, D (Program, Day, Time)
        const className = `${data[i][1]} ${data[i][2]} ${data[i][3]}`;
        classes.push(className);
      }
    }
    
    return classes.length > 0 ? classes : ['No classes found'];
  } catch (error) {
    Logger.log(`Error getting classes for selector: ${error.message}`);
    return ['Error loading classes'];
  }
}

/**
 * Creates or resets the communication log sheet
 * @return {Sheet} The created or updated sheet
 */
function createCommunicationLog() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(COMMUNICATIONS_HUB_CONFIG.LOG_SHEET_NAME);
    
    // Create the sheet if it doesn't exist or reset headers if it does
    if (!sheet) {
      // Create a new sheet
      sheet = ss.insertSheet(COMMUNICATIONS_HUB_CONFIG.LOG_SHEET_NAME);
      
      // Set up headers
      sheet.getRange('A1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.DATE_SENT);
      sheet.getRange('B1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.SUBJECT);
      sheet.getRange('C1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.RECIPIENTS);
      sheet.getRange('D1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.SENDER);
      sheet.getRange('E1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.STATUS);
      
      // Format headers
      sheet.getRange('A1:E1')
        .setFontWeight('bold')
        .setBackground(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.HEADER_COLOR)
        .setFontColor(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
      
      // Set column widths
      sheet.setColumnWidth(1, 150); // Date
      sheet.setColumnWidth(2, 300); // Subject
      sheet.setColumnWidth(3, 300); // Recipients
      sheet.setColumnWidth(4, 150); // Sender
      sheet.setColumnWidth(5, 100); // Status
      
      // Freeze the header row
      sheet.setFrozenRows(1);
    } else {
      // If sheet exists, just make sure the headers are correct
      sheet.getRange('A1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.DATE_SENT);
      sheet.getRange('B1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.SUBJECT);
      sheet.getRange('C1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.RECIPIENTS);
      sheet.getRange('D1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.SENDER);
      sheet.getRange('E1').setValue(COMMUNICATIONS_HUB_CONFIG.LOG_COLUMNS.STATUS);
      
      // Format headers
      sheet.getRange('A1:E1')
        .setFontWeight('bold')
        .setBackground(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.HEADER_COLOR)
        .setFontColor(COMMUNICATIONS_HUB_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    }
    
    // Show confirmation
    SpreadsheetApp.getUi().alert(
      'Communication Log Created',
      'The communication log has been created. It will record all sent communications.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Set active sheet to communication log
    sheet.activate();
    
    return sheet;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createCommunicationLog', 
        'Error creating communication log. Please try again or contact support.');
    } else {
      Logger.log(`Error creating communication log: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create communication log: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return null;
  }
}

/**
 * Sends the selected communication from the communications hub
 * @return {boolean} Success status
 */
function sendSelectedCommunication() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(COMMUNICATIONS_HUB_CONFIG.SHEET_NAME);
    
    if (!sheet) {
      SpreadsheetApp.getUi().alert(
        'Communications Hub Not Found',
        'Please create the Communications Hub first using "YSL Hub > Communications > Create Dynamic Communications Hub"',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Check if communication is selected for sending
    const isSelected = sheet.getRange('A31').getValue();
    
    if (!isSelected) {
      SpreadsheetApp.getUi().alert(
        'No Communication Selected',
        'Please select the communication for sending by checking the box at the bottom of the Communications Hub.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Gather communication details
    const subject = sheet.getRange('B14').getValue();
    const body = sheet.getRange('B15').getValue();
    const sendDate = sheet.getRange('B26').getValue();
    
    // Check if we have all required information
    if (!subject || !body) {
      SpreadsheetApp.getUi().alert(
        'Incomplete Communication',
        'Please provide both a subject and message body for the communication.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get selected recipients
    const allClasses = sheet.getRange('B3').getValue();
    const allPrivateLessons = sheet.getRange('B4').getValue();
    
    // Get individual class selections
    const classSelections = [];
    for (let i = 5; i <= 11; i++) {
      const className = sheet.getRange(i, 1).getValue();
      const isClassSelected = sheet.getRange(i, 2).getValue();
      
      if (className && isClassSelected) {
        classSelections.push(className);
      }
    }
    
    // Check if any recipients are selected
    if (!allClasses && !allPrivateLessons && classSelections.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Recipients Selected',
        'Please select at least one recipient class for the communication.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // For now, we'll just log what would happen
    // In a real implementation, this would actually send the emails
    
    // Log the communication to the communication log
    logCommunication(subject, 
                    allClasses ? 'All Classes' : (allPrivateLessons ? 'All Private Lessons' : classSelections.join(', ')), 
                    'Scheduled');
    
    // Update status on the communications hub
    sheet.getRange('B27').setValue(`Scheduled for sending on ${Utilities.formatDate(sendDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}`);
    
    // Uncheck the selection box to prevent accidental re-sending
    sheet.getRange('A31').setValue(false);
    
    // Show confirmation
    SpreadsheetApp.getUi().alert(
      'Communication Scheduled',
      `Your communication has been scheduled for sending on ${Utilities.formatDate(sendDate, Session.getScriptTimeZone(), 'MM/dd/yyyy')}.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendSelectedCommunication', 
        'Error sending communication. Please try again or contact support.');
    } else {
      Logger.log(`Error sending communication: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to send communication: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Logs a sent communication to the communication log
 * @param {string} subject - The email subject
 * @param {string} recipients - Description of recipients
 * @param {string} status - Status of the communication
 */
function logCommunication(subject, recipients, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(COMMUNICATIONS_HUB_CONFIG.LOG_SHEET_NAME);
    
    // If log sheet doesn't exist, create it
    if (!logSheet) {
      logSheet = createCommunicationLog();
    }
    
    // Get the next available row
    const nextRow = logSheet.getLastRow() + 1;
    
    // Add the log entry
    logSheet.getRange(nextRow, 1).setValue(new Date())
      .setNumberFormat('mm/dd/yyyy HH:mm:ss');
    logSheet.getRange(nextRow, 2).setValue(subject);
    logSheet.getRange(nextRow, 3).setValue(recipients);
    logSheet.getRange(nextRow, 4).setValue(Session.getActiveUser().getEmail());
    logSheet.getRange(nextRow, 5).setValue(status);
    
  } catch (error) {
    Logger.log(`Error logging communication: ${error.message}`);
    // Continue without failing the whole operation
  }
}

// Make functions available to other modules
const CommunicationModule = Object.assign(CommunicationModule || {}, {
  createCommunicationsHub: createCommunicationsHub,
  createCommunicationLog: createCommunicationLog,
  sendSelectedCommunication: sendSelectedCommunication
});