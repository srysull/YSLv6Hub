/**
 * YSL Hub v2 Global Functions
 * 
 * This module provides common functions and event handlers accessible
 * to all other modules in the system.
 * 
 * @author PenBay YMCA
 * @version 2.0
 * @date 2025-04-27
 */

/**
 * Handles when the spreadsheet is opened.
 * Sets up the menu and initializes the system if needed.
 */
function onOpen() {
  try {
    // Initialize error handling first for proper logging
    if (ErrorHandling && typeof ErrorHandling.initializeErrorHandling === 'function') {
      ErrorHandling.initializeErrorHandling();
    }
    
    // Log that the spreadsheet was opened
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Spreadsheet opened', 'INFO', 'onOpen');
    }
    
    // Initialize version control
    if (VersionControl && typeof VersionControl.initializeVersionControl === 'function') {
      VersionControl.initializeVersionControl();
    }
    
    // Add the menu
    AdministrativeModule.createMenu();
  } catch (error) {
    // Log error, using native Logger as fallback since error handling might not be initialized
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'onOpen', 'Error during system initialization.');
    } else {
      Logger.log(`Error in onOpen: ${error.message}`);
      SpreadsheetApp.getUi().alert('Initialization Error', 
        `An error occurred during initialization: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Handles edits to the spreadsheet and performs appropriate actions.
 * This function is triggered automatically when a cell is edited.
 * 
 * @param {Object} e - The edit event object
 */
function onEdit(e) {
  try {
    // Get the edited range information
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    const value = e.value;
    
    // Log the edit event
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Cell edited: ${sheetName} (${row}, ${col}) = ${value}`, 'DEBUG', 'onEdit');
    }
    
    // Handle edits in Classes sheet
    if (sheetName === 'Classes') {
      handleClassesSheetEdit(sheet, row, col, value);
    }
    // Handle edits in Announcements sheet
    else if (sheetName === 'Announcements') {
      handleAnnouncementsSheetEdit(sheet, row, col, value);
    }
  } catch (error) {
    // Log the error but don't show alert to avoid disrupting user
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error in onEdit: ${error.message}`, 'ERROR', 'onEdit');
    } else {
      Logger.log(`Error in onEdit: ${error.message}`);
    }
  }
}

/**
 * Handles edits in the Classes sheet.
 * 
 * @param {Sheet} sheet - The Classes sheet
 * @param {number} row - The edited row
 * @param {number} col - The edited column
 * @param {string} value - The new cell value
 */
function handleClassesSheetEdit(sheet, row, col, value) {
  // Check if the edit is in the "Select Class" column (column 1)
  if (col === 1 && row > 1 && value === 'Select') {
    try {
      // Get class information
      const classInfo = {
        program: sheet.getRange(row, 2).getValue(),
        day: sheet.getRange(row, 3).getValue(),
        time: sheet.getRange(row, 4).getValue(),
        location: sheet.getRange(row, 5).getValue(),
        count: sheet.getRange(row, 6).getValue(),
        instructor: sheet.getRange(row, 7).getValue()
      };
      
      // Check if instructor is set
      if (!classInfo.instructor) {
        const ui = SpreadsheetApp.getUi();
        const result = ui.alert(
          'Missing Instructor',
          'No instructor is set for this class. Would you like to add an instructor now?',
          ui.ButtonSet.YES_NO
        );
        
        if (result === ui.Button.YES) {
          const response = ui.prompt(
            'Add Instructor',
            'Enter the instructor name:',
            ui.ButtonSet.OK_CANCEL
          );
          
          if (response.getSelectedButton() === ui.Button.OK) {
            const instructorName = response.getResponseText().trim();
            if (instructorName) {
              sheet.getRange(row, 7).setValue(instructorName);
              classInfo.instructor = instructorName;
            }
          }
        }
      }
      
      // Offer to create instructor sheet
      const ui = SpreadsheetApp.getUi();
      const result = ui.alert(
        'Create Instructor Sheet',
        `Would you like to create an instructor sheet for ${classInfo.program} (${classInfo.day}, ${classInfo.time})?`,
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        InstructorResourceModule.createInstructorSheet(classInfo);
      }
    } catch (error) {
      // Log error with proper error handling
      if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
        ErrorHandling.handleError(error, 'handleClassesSheetEdit', 
          'Error creating instructor sheet. Please try again or contact support.');
      } else {
        Logger.log(`Error handling Classes sheet edit: ${error.message}`);
      }
    }
  }
}

/**
 * Handles edits in the Announcements sheet.
 * 
 * @param {Sheet} sheet - The Announcements sheet
 * @param {number} row - The edited row
 * @param {number} col - The edited column
 * @param {string} value - The new cell value
 */
function handleAnnouncementsSheetEdit(sheet, row, col, value) {
  // Check if the edit is in the Status column (column 9)
  if (col === 9 && row > 1 && value === 'Ready') {
    try {
      // Get announcement data
      const announcement = {
        classId: sheet.getRange(row, 1).getValue(),
        subject: sheet.getRange(row, 7).getValue(),
        message: sheet.getRange(row, 8).getValue()
      };
      
      // Validate announcement
      if (!announcement.subject || !announcement.message) {
        const ui = SpreadsheetApp.getUi();
        ui.alert(
          'Invalid Announcement',
          'Subject and message are required for announcements. Please complete these fields before marking as Ready.',
          ui.ButtonSet.OK
        );
        
        // Reset status to Draft
        sheet.getRange(row, 9).setValue('Draft');
      }
    } catch (error) {
      // Log error with proper error handling
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Error handling Announcements sheet edit: ${error.message}`, 'ERROR', 'handleAnnouncementsSheetEdit');
      } else {
        Logger.log(`Error handling Announcements sheet edit: ${error.message}`);
      }
    }
  }
}

/**
 * Safely extract an ID from a Google Drive URL.
 * This function handles various URL formats and provides better validation.
 * 
 * @param {string} url - The Google Drive URL
 * @return {string|null} The extracted ID, or null if not found
 */
function extractIdFromUrl(url) {
  if (!url) return null;
  
  // Log the URL for debugging
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage(`Extracting ID from URL: ${url}`, 'DEBUG', 'extractIdFromUrl');
  }
  
  // Check if URL is already an ID (just alphanumeric characters and dashes)
  if (/^[a-zA-Z0-9_-]+$/.test(url)) {
    return url;
  }
  
  // Extract ID from various URL formats
  const patterns = [
    /\/folders\/([a-zA-Z0-9_-]+)/,         // Drive folder URL
    /\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/, // Spreadsheet URL
    /id=([a-zA-Z0-9_-]+)/,                 // URL parameter format
    /\/file\/d\/([a-zA-Z0-9_-]+)/,         // Direct file URL
    /docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/ // Google Doc URL
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  // Additional validation for embedded IDs
  if (url.includes('1vlR8WwEyLWOuO-JUzCzrdLikv-hVVld4')) {
    return '1vlR8WwEyLWOuO-JUzCzrdLikv-hVVld4';
  }
  
  // If no match found in common patterns, check if the entire URL might be a valid ID
  if (/[a-zA-Z0-9_-]{25,45}/.test(url)) {
    const idMatch = url.match(/([a-zA-Z0-9_-]{25,45})/);
    if (idMatch && idMatch[1]) {
      return idMatch[1];
    }
  }
  
  return null;
}

/**
 * Safe access to Google Drive folder by ID with proper error handling
 * 
 * @param {string} folderId - The folder ID to access
 * @return {Folder|null} The folder object or null if not found
 */
function safeGetFolderById(folderId) {
  if (!folderId) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Attempted to access folder with null/undefined ID', 'ERROR', 'safeGetFolderById');
    }
    return null;
  }
  
  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error accessing folder with ID ${folderId}: ${error.message}`, 'ERROR', 'safeGetFolderById');
    } else {
      Logger.log(`Error accessing folder with ID ${folderId}: ${error.message}`);
    }
    return null;
  }
}

/**
 * Safe access to Google Drive file by ID with proper error handling
 * 
 * @param {string} fileId - The file ID to access
 * @return {File|null} The file object or null if not found
 */
function safeGetFileById(fileId) {
  if (!fileId) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Attempted to access file with null/undefined ID', 'ERROR', 'safeGetFileById');
    }
    return null;
  }
  
  try {
    return DriveApp.getFileById(fileId);
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error accessing file with ID ${fileId}: ${error.message}`, 'ERROR', 'safeGetFileById');
    } else {
      Logger.log(`Error accessing file with ID ${fileId}: ${error.message}`);
    }
    return null;
  }
}

/**
 * Safe access to Spreadsheet by ID with proper error handling
 * 
 * @param {string} spreadsheetId - The spreadsheet ID to access
 * @return {Spreadsheet|null} The spreadsheet object or null if not found
 */
function safeGetSpreadsheetById(spreadsheetId) {
  if (!spreadsheetId) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Attempted to access spreadsheet with null/undefined ID', 'ERROR', 'safeGetSpreadsheetById');
    }
    return null;
  }
  
  try {
    return SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error accessing spreadsheet with ID ${spreadsheetId}: ${error.message}`, 'ERROR', 'safeGetSpreadsheetById');
    } else {
      Logger.log(`Error accessing spreadsheet with ID ${spreadsheetId}: ${error.message}`);
    }
    return null;
  }
}

/**
 * Formats a student name to show only first name and first initial of last name
 * 
 * @param {string} fullName - The student's full name
 * @return {string} Formatted name (e.g. "John D.")
 */
function formatStudentName(fullName) {
  if (!fullName) return '';
  
  const nameParts = fullName.trim().split(' ');
  if (nameParts.length === 1) return nameParts[0];
  
  const firstName = nameParts[0];
  const lastInitial = nameParts[nameParts.length - 1].charAt(0);
  
  return `${firstName} ${lastInitial}.`;
}

/**
 * Finds a column index by header name, case-insensitive, with fallback options
 * 
 * @param {Array} headers - Array of header names
 * @param {Array|string} possibleNames - Possible name(s) for the column
 * @return {number} The index of the found column, or -1 if not found
 */
function findColumnIndex(headers, possibleNames) {
  // Convert possibleNames to an array if it's a string
  const nameOptions = Array.isArray(possibleNames) ? possibleNames : [possibleNames];
  
  // First try exact matches (case-insensitive)
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] ? headers[i].toString().toLowerCase() : '';
    
    for (const name of nameOptions) {
      if (header === name.toLowerCase()) {
        return i; // Exact match
      }
    }
  }
  
  // Then try partial matches (case-insensitive)
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] ? headers[i].toString().toLowerCase() : '';
    
    for (const name of nameOptions) {
      if (header.includes(name.toLowerCase()) || name.toLowerCase().includes(header)) {
        return i; // Partial match
      }
    }
  }
  
  return -1; // No match found
}

/**
 * Safely gets a property with proper error handling
 * 
 * @param {string} key - The property key to get
 * @param {string} defaultValue - Default value if property not found
 * @return {string} The property value or default
 */
function safeGetProperty(key, defaultValue = '') {
  try {
    const value = PropertiesService.getScriptProperties().getProperty(key);
    return value !== null ? value : defaultValue;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting property ${key}: ${error.message}`, 'ERROR', 'safeGetProperty');
    } else {
      Logger.log(`Error getting property ${key}: ${error.message}`);
    }
    return defaultValue;
  }
}

/**
 * Safely sets a property with proper error handling
 * 
 * @param {string} key - The property key to set
 * @param {string} value - The value to set
 * @return {boolean} Success status
 */
function safeSetProperty(key, value) {
  try {
    PropertiesService.getScriptProperties().setProperty(key, value);
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error setting property ${key}: ${error.message}`, 'ERROR', 'safeSetProperty');
    } else {
      Logger.log(`Error setting property ${key}: ${error.message}`);
    }
    return false;
  }
}

// Make functions available to other modules
const GlobalFunctions = {
  onOpen: onOpen,
  onEdit: onEdit,
  extractIdFromUrl: extractIdFromUrl,
  safeGetFolderById: safeGetFolderById,
  safeGetFileById: safeGetFileById,
  safeGetSpreadsheetById: safeGetSpreadsheetById,
  formatStudentName: formatStudentName,
  findColumnIndex: findColumnIndex,
  safeGetProperty: safeGetProperty,
  safeSetProperty: safeSetProperty
};