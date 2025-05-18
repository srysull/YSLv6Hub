/**
 * YSL Hub v2 Global Functions
 * 
 * This module provides common functions and event handlers accessible
 * to all other modules in the system.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-14
 */

/**
 * Handles when the spreadsheet is opened.
 * This function is DEPRECATED - use the onOpen in 00_TriggerFunctions.ts instead.
 * This duplicate function is kept for compatibility but now forwards to the main trigger function.
 */
function onOpen_Old() {
  // This function is intentionally renamed to avoid trigger confusion
  // The actual onOpen functionality has been moved to 00_TriggerFunctions.ts
  
  // Log that this deprecated function was called
  Logger.log('WARNING: The onOpen function in 01_Globals.ts was called, but this is deprecated.');
  Logger.log('Use the onOpen function in 00_TriggerFunctions.ts instead.');
  
  try {
    // Initialize error handling first for proper logging
    if (ErrorHandling && typeof ErrorHandling.initializeErrorHandling === 'function') {
      ErrorHandling.initializeErrorHandling();
    }
    
    // Log that the spreadsheet was opened
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Spreadsheet opened via deprecated onOpen function', 'WARNING', 'onOpen_Old');
    }
    
    // Initialize version control
    if (VersionControl && typeof VersionControl.initializeVersionControl === 'function') {
      VersionControl.initializeVersionControl();
    }
    
    // Force properties to true to ensure menu appears
    PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    // Call the emergency menu fix function from MenuFix.gs if available
    if (typeof createFixedMenu === 'function') {
      createFixedMenu();
      return;
    }
    
    // Add the menu as fallback
    if (typeof AdministrativeModule !== 'undefined' && 
        typeof AdministrativeModule.createMenu === 'function') {
      AdministrativeModule.createMenu();
    }
  } catch (error) {
    // Log error, using native Logger as fallback since error handling might not be initialized
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'onOpen_Old', 'Error during system initialization.');
    } else {
      Logger.log(`Error in onOpen_Old: ${error.message}`);
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
 * @param e - The edit event object
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
    // Handle edits in Group Lesson Tracker sheet
    else if (sheetName === 'Group Lesson Tracker') {
      if (row === 2 && col === 1) {
        // Cell A2 (dropdown) in Group Lesson Tracker was edited
        handleGroupLessonTrackerDropdownChange(sheet, value);
      } else if (row === 4 && col === 1) {
        // Cell A4 (sync button) was clicked
        if (typeof syncSwimmerData === 'function') {
          // Call the global sync function
          syncSwimmerData();
        } else {
          // Fallback to the direct function
          syncStudentDataWithSwimmerSkills(sheet);
        }
        
        // Reset the cell text
        Utilities.sleep(500);
        sheet.getRange('A4').setValue('Click here to SYNC DATA');
      }
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
 * @param sheet - The Classes sheet
 * @param row - The edited row
 * @param col - The edited column
 * @param value - The new cell value for the checkbox
 */
function handleClassesSheetEdit(sheet, row, col, value) {
  // Check if the edit is in the checkbox column (column 1)
  if (col === 1 && row > 1 && value === true) {
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
 * @param sheet - The Announcements sheet
 * @param row - The edited row
 * @param col - The edited column
 * @param value - The new cell value
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
 * @param url - The Google Drive URL
 * @return The extracted ID, or null if not found
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
 * @param folderId - The folder ID to access
 * @return The folder object or null if not found
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
 * @param fileId - The file ID to access
 * @return The file object or null if not found
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
 * @param spreadsheetId - The spreadsheet ID to access
 * @return The spreadsheet object or null if not found
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
 * @param fullName - The student's full name
 * @return Formatted name (e.g. "John D.")
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
 * @param headers - Array of header names
 * @param possibleNames - Possible name(s) for the column
 * @return The index of the found column, or -1 if not found
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
 * @param key - The property key to get
 * @param defaultValue - Default value if property not found
 * @return The property value or default
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
 * @param key - The property key to set
 * @param value - The value to set
 * @return Success status
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
/**
 * Handles changes to the dropdown in the Group Lesson Tracker sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param selectedClass - The selected class value
 */
function handleGroupLessonTrackerDropdownChange(sheet, selectedClass) {
  try {
    // Log the event
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Group Lesson Tracker dropdown changed to: ${selectedClass}`, 'INFO', 'handleGroupLessonTrackerDropdownChange');
    }
    
    // If no class is selected or it's a placeholder, do nothing
    if (!selectedClass || 
        selectedClass === 'Select a class...' || 
        selectedClass === 'No classes available' || 
        selectedClass === 'Error loading classes') {
      return;
    }
    
    // Clear existing data before populating with new class data
    clearExistingTrackerData(sheet);
    
    // Get the Daxko sheet to read student data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName('Daxko');
    
    if (!daxkoSheet) {
      Logger.log('Daxko sheet not found');
      return;
    }
    
    // Get all data from Daxko sheet
    const daxkoData = daxkoSheet.getDataRange().getValues();
    
    // Find the row that matches the selected class
    let matchingRow = null;
    let programName = '';
    
    for (let i = 1; i < daxkoData.length; i++) {
      // Recreate the concatenated class name (W and X columns)
      const classW = daxkoData[i][22]; // Column W (0-indexed, so 22 = column W)
      const classX = daxkoData[i][23]; // Column X
      
      // Only concatenate if both values exist
      let concatenatedClass = '';
      if (classW) concatenatedClass += classW.toString().trim();
      if (classX) concatenatedClass += ' ' + classX.toString().trim();
      concatenatedClass = concatenatedClass.trim();
      
      if (concatenatedClass === selectedClass) {
        matchingRow = i;
        programName = classW ? classW.toString().trim() : '';
        break;
      }
    }
    
    if (matchingRow === null) {
      Logger.log(`No matching class found for: ${selectedClass}`);
      return;
    }
    
    // Get program abbreviation
    const programAbbreviation = getProgramAbbreviation(programName);
    
    // 1. Populate student names (B4:X4)
    const studentsData = populateStudentNames(sheet, daxkoData, matchingRow);
    
    // 2. Populate class dates (A5-A12)
    populateClassDates(sheet, daxkoData, matchingRow);
    
    // 3. Populate class skills (A14-A29)
    const skillsMapping = populateClassSkills(sheet, programAbbreviation);
    
    // 4. Populate beginning skills for each student
    if (studentsData && studentsData.length > 0 && skillsMapping) {
      populateStudentSkills(sheet, studentsData, skillsMapping);
    }
    
    // Log successful update
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Group Lesson Tracker sheet successfully updated for: ${selectedClass}`, 'INFO', 'handleGroupLessonTrackerDropdownChange');
    }
  } catch (error) {
    // Log the error
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error updating Group Lesson Tracker: ${error.message}`, 'ERROR', 'handleGroupLessonTrackerDropdownChange');
    } else {
      Logger.log(`Error updating Group Lesson Tracker: ${error.message}`);
    }
  }
}

/**
 * Clears existing data in the Group Lesson Tracker sheet
 * Clears ranges B5:ZZ12 (attendance), B14:ZZ29 (skills 1), and B31:ZZ40 (skills 2)
 * 
 * @param sheet - The Group Lesson Tracker sheet
 */
function clearExistingTrackerData(sheet) {
  try {
    Logger.log('Clearing existing data from Group Lesson Tracker');
    
    // Get the last column of the sheet
    const lastColumn = sheet.getLastColumn();
    
    // Convert to column letter notation
    const lastColLetter = columnToLetter(lastColumn);
    
    // Clear attendance data (B5:ZZ12)
    const attendanceRange = sheet.getRange(`B5:${lastColLetter}12`);
    attendanceRange.clearContent();
    attendanceRange.clearFormat();
    Logger.log(`Cleared attendance data (B5:${lastColLetter}12)`);
    
    // Clear skills section 1 (B14:ZZ30) - extended to include row 30
    const skills1Range = sheet.getRange(`B14:${lastColLetter}30`);
    skills1Range.clearContent();
    skills1Range.clearFormat();
    Logger.log(`Cleared skills section 1 (B14:${lastColLetter}30)`);
    
    // Clear skills section 2 (B32:ZZ41) - updated to 32-41 range
    const skills2Range = sheet.getRange(`B32:${lastColLetter}41`);
    skills2Range.clearContent();
    skills2Range.clearFormat();
    Logger.log(`Cleared skills section 2 (B32:${lastColLetter}41)`);
    
    // Re-apply light gray background to alternating cells
    applyAlternatingBackground(sheet);
    
    Logger.log('Successfully cleared existing data from Group Lesson Tracker');
  } catch (error) {
    Logger.log(`Error clearing existing data: ${error.message}`);
  }
}

/**
 * Apply alternating light gray background to skill cells
 * 
 * @param sheet - The Group Lesson Tracker sheet
 */
function applyAlternatingBackground(sheet) {
  try {
    // Get the last column of the sheet
    const lastColumn = sheet.getLastColumn();
    
    // Define the light background color
    const lightGray = '#F9F9F9';
    
    // Apply background to every other row in the skills sections
    for (let row = 14; row <= 30; row += 2) {
      sheet.getRange(row, 2, 1, lastColumn - 1).setBackground(lightGray);
    }
    
    for (let row = 32; row <= 41; row += 2) {
      sheet.getRange(row, 2, 1, lastColumn - 1).setBackground(lightGray);
    }
  } catch (error) {
    Logger.log(`Error applying alternating background: ${error.message}`);
  }
}

/**
 * Converts a column number to letter notation (e.g., 1 -> A, 27 -> AA)
 * 
 * @param column - The column number
 * @return The column letter notation
 */
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * Extracts program abbreviation from program name
 * Takes first letter of first word and first character of second word
 * 
 * @param programName - The program name (e.g., "Stage 1 - Water Acclimation")
 * @returns The program abbreviation (e.g., "S1")
 */
function getProgramAbbreviation(programName) {
  try {
    if (!programName) return '';
    
    const words = programName.split(' ');
    if (words.length < 2) return programName.charAt(0);
    
    // First letter of first word
    const firstLetter = words[0].charAt(0).toUpperCase();
    
    // First character of second word (might be a number or letter)
    const secondChar = words[1].charAt(0);
    
    return firstLetter + secondChar;
  } catch (error) {
    Logger.log(`Error getting program abbreviation: ${error.message}`);
    return '';
  }
}

/**
 * Populates student names in the Group Lesson Tracker sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param daxkoData - All data from the Daxko sheet
 * @param matchingRow - The row index matching the selected class
 */
/**
 * Populates student names in the Group Lesson Tracker sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param daxkoData - All data from the Daxko sheet
 * @param matchingRow - The row index matching the selected class
 * @returns Array of student data objects
 */
function populateStudentNames(sheet, daxkoData, matchingRow) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Get student data - we need to find all students with the same class
    const matchingClass = {
      w: daxkoData[matchingRow][22], // Column W
      x: daxkoData[matchingRow][23], // Column X
      z: daxkoData[matchingRow][25]  // Column Z (day)
    };
    
    // Collect all students in this class
    const students = [];
    for (let i = 1; i < daxkoData.length; i++) {
      // Check if this row matches our class
      if (daxkoData[i][22] === matchingClass.w && 
          daxkoData[i][23] === matchingClass.x && 
          daxkoData[i][25] === matchingClass.z) {
        
        // Get student first and last name from columns C and D
        const firstName = daxkoData[i][2] || ''; // Column C
        const lastName = daxkoData[i][3] || '';  // Column D
        
        // Only add if we have a name
        if (firstName || lastName) {
          students.push({
            firstName: firstName,
            lastName: lastName,
            fullName: `${firstName} ${lastName}`.trim()
          });
        }
      }
    }
    
    // Define the columns for student names
    // These are the columns where student names will appear, created in pairs
    // Each student takes two columns for B/E assessment
    const studentNameColumns = ['B', 'D', 'F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'V', 'X', 
                               'Z', 'AB', 'AD', 'AF', 'AH', 'AJ', 'AL', 'AN', 'AP', 'AR', 'AT', 'AV'];
    
    // Track the current sheet width
    const currentLastColumn = sheet.getLastColumn();
    
    // Clear existing student names in the original template (first 12 student positions)
    for (let i = 0; i < Math.min(studentNameColumns.length, 12); i++) {
      sheet.getRange(studentNameColumns[i] + '4').setValue('');
    }
    
    // If we previously expanded beyond Y column (column 25) but don't need it now
    // or if we had previously expanded more than we need now, truncate the sheet
    if (currentLastColumn > 25) {
      if (students.length <= 12) {
        // Truncate back to original size
        truncateSheetToOriginalSize(sheet);
      } else {
        // Calculate the exact columns needed (column A + 2 columns per student)
        const exactColumnsNeeded = 1 + (students.length * 2);
        
        // If we have more columns than we need, truncate
        if (currentLastColumn > exactColumnsNeeded) {
          truncateSheetToSize(sheet, exactColumnsNeeded);
        }
      }
    }
    
    // Now we can expand if needed - this will now add only the exact columns needed
    if (students.length > 12) {
      expandSheetForStudents(sheet, students.length);
      
      // Clear any expanded student name cells from previous selections 
      // We do this after expansion to make sure the cells exist
      for (let i = 12; i < Math.min(studentNameColumns.length, students.length); i++) {
        try {
          sheet.getRange(studentNameColumns[i] + '4').setValue('');
        } catch (e) {
          // Cell might not exist, just continue
        }
      }
    }
    
    // Populate student names
    for (let i = 0; i < students.length; i++) {
      // Get the appropriate cell reference for this student
      const cellRef = studentNameColumns[i] + '4';
      
      // Set the student name
      try {
        sheet.getRange(cellRef).setValue(students[i].fullName);
        
        // Add column information to student object for later use
        students[i].beginningColumn = studentNameColumns[i];
        
        // Calculate the end column (always next column)
        const endColIndex = studentNameColumns.indexOf(studentNameColumns[i]) + 1;
        if (endColIndex < studentNameColumns.length) {
          students[i].endingColumn = studentNameColumns[endColIndex];
        }
      } catch (e) {
        // Handle any issues with setting values
        Logger.log(`Error setting student name in cell ${cellRef}: ${e.message}`);
      }
    }
    
    // Log the number of students found
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Populated ${students.length} student names in Group Lesson Tracker`, 'INFO', 'populateStudentNames');
    }
    
    // Return student data for further processing
    return students;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error populating student names: ${error.message}`, 'ERROR', 'populateStudentNames');
    } else {
      Logger.log(`Error populating student names: ${error.message}`);
    }
    return [];
  }
}

/**
 * Truncates the sheet back to its original size (column Y = 25)
 * 
 * @param sheet - The Group Lesson Tracker sheet
 */
function truncateSheetToOriginalSize(sheet) {
  try {
    const currentLastColumn = sheet.getLastColumn();
    
    // Only proceed if we have more than the original 25 columns (A-Y)
    if (currentLastColumn <= 25) {
      return;
    }
    
    // Calculate how many columns to delete (columns beyond Y)
    const columnsToDelete = currentLastColumn - 25;
    
    // Delete the extra columns
    sheet.deleteColumns(26, columnsToDelete);
    
    // Log the truncation
    Logger.log(`Truncated sheet back to original size (deleted ${columnsToDelete} columns)`);
    
    // Re-merge the header row
    const titleRow = sheet.getRange(1, 1, 1, 25);
    if (!titleRow.isPartOfMerge()) {
      titleRow.merge();
    }
    
    const dropdownRow = sheet.getRange(2, 1, 1, 25);
    if (!dropdownRow.isPartOfMerge()) {
      dropdownRow.merge();
    }
    
    const attendanceHeaderRow = sheet.getRange(3, 1, 1, 25);
    if (!attendanceHeaderRow.isPartOfMerge()) {
      attendanceHeaderRow.merge();
    }
  } catch (error) {
    // Log the error but don't throw
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error truncating sheet: ${error.message}`, 'ERROR', 'truncateSheetToOriginalSize');
    } else {
      Logger.log(`Error truncating sheet: ${error.message}`);
    }
  }
}

/**
 * Truncates the sheet to a specific number of columns
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param targetColumns - The number of columns to keep
 */
function truncateSheetToSize(sheet, targetColumns) {
  try {
    const currentLastColumn = sheet.getLastColumn();
    
    // Only proceed if we have more columns than needed
    if (currentLastColumn <= targetColumns) {
      return;
    }
    
    // Calculate how many columns to delete
    const columnsToDelete = currentLastColumn - targetColumns;
    
    // Delete the extra columns
    sheet.deleteColumns(targetColumns + 1, columnsToDelete);
    
    // Log the truncation
    Logger.log(`Truncated sheet to ${targetColumns} columns (deleted ${columnsToDelete} columns)`);
    
    // Re-merge the header row
    const titleRow = sheet.getRange(1, 1, 1, targetColumns);
    if (!titleRow.isPartOfMerge()) {
      titleRow.merge();
    }
    
    const dropdownRow = sheet.getRange(2, 1, 1, targetColumns);
    if (!dropdownRow.isPartOfMerge()) {
      dropdownRow.merge();
    }
    
    const attendanceHeaderRow = sheet.getRange(3, 1, 1, targetColumns);
    if (!attendanceHeaderRow.isPartOfMerge()) {
      attendanceHeaderRow.merge();
    }
  } catch (error) {
    // Log the error but don't throw
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error truncating sheet: ${error.message}`, 'ERROR', 'truncateSheetToSize');
    } else {
      Logger.log(`Error truncating sheet: ${error.message}`);
    }
  }
}

/**
 * Expands the sheet to accommodate more students if needed
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param studentCount - The number of students to accommodate
 */
function expandSheetForStudents(sheet, studentCount) {
  try {
    // Default template is designed for 12 students (24 columns: B-Y)
    // If we have more than 12 students, we need to expand
    
    // Only proceed if we have more than 12 students
    if (studentCount <= 12) {
      return;
    }
    
    // Calculate number of additional student columns needed
    // Each student requires 2 columns (one for beginning, one for ending)
    // We only add the exact number of columns needed for students beyond 12
    const additionalStudentsNeeded = studentCount - 12; // Only add what's needed
    const additionalColumnsNeeded = additionalStudentsNeeded * 2;
    
    // Check current sheet width
    const lastColumn = sheet.getLastColumn();
    
    // Y column is 25, so we need to add columns if lastColumn < 25 + additionalColumnsNeeded
    const targetLastColumn = 25 + additionalColumnsNeeded;
    
    // Only add columns if we don't already have enough
    if (lastColumn < targetLastColumn) {
      // Add the needed columns
      sheet.insertColumnsAfter(lastColumn, targetLastColumn - lastColumn);
      
      // Log the expansion
      Logger.log(`Expanded sheet to accommodate ${studentCount} students (added ${targetLastColumn - lastColumn} columns)`);
    }
    
    // Now format the additional columns similar to the existing ones
    for (let i = 26; i <= targetLastColumn; i++) {
      // Set column width to 40px like the original student columns
      sheet.setColumnWidth(i, 40);
    }
    
    // Define student column pairs to merge - we're extending beyond Y column
    // These are indexed from 0, so column 25 is Z
    const baseColumns = 25; // Z is the first column after the original template
    
    // Create column pairs for the additional students
    for (let i = 0; i < additionalStudentsNeeded; i++) {
      // Calculate the start column for this student pair
      const startCol = baseColumns + (i * 2);
      
      // Create the student name header cell and merge
      const nameRange = sheet.getRange(4, startCol, 1, 2);
      nameRange.setValue(`Student ${i+13}`) // Student numbers start at 13 for the expanded section
        .setFontWeight('bold')
        .setBackground('#F3F3F3')
        .setHorizontalAlignment('center');
      
      // Merge the name cell
      if (!nameRange.isPartOfMerge()) {
        nameRange.merge();
      }
      
      // For each row in the attendance section (rows 5-12)
      for (let row = 5; row <= 12; row++) {
        const attendanceRange = sheet.getRange(row, startCol, 1, 2);
        attendanceRange.setBackground('#F9F9F9');
        
        // Merge the attendance cell
        if (!attendanceRange.isPartOfMerge()) {
          attendanceRange.merge();
        }
      }
      
      // Set Beginning/End headers for row 13 (skills section header)
      sheet.getRange(13, startCol).setValue('B') // B for Beginning
        .setFontWeight('bold')
        .setBackground('#F3F3F3')
        .setHorizontalAlignment('center');
      
      sheet.getRange(13, startCol + 1).setValue('E') // E for End
        .setFontWeight('bold')
        .setBackground('#F3F3F3')
        .setHorizontalAlignment('center');
      
      // For each row in the skills sections (rows 14-43)
      for (let row = 14; row <= 43; row++) {
        // Format B cell
        sheet.getRange(row, startCol)
          .setBackground('#F9F9F9');
        
        // Format E cell
        sheet.getRange(row, startCol + 1)
          .setBackground('#F9F9F9');
      }
      
      // Add B/E headers for row 30 (second B/E header row)
      sheet.getRange(30, startCol).setValue('B')
        .setFontWeight('bold')
        .setBackground('#F3F3F3')
        .setHorizontalAlignment('center');
      
      sheet.getRange(30, startCol + 1).setValue('E')
        .setFontWeight('bold')
        .setBackground('#F3F3F3')
        .setHorizontalAlignment('center');
    }
    
    // Ensure the header row spans across all columns
    const titleRow = sheet.getRange(1, 1, 1, targetLastColumn);
    if (!titleRow.isPartOfMerge()) {
      titleRow.merge();
    }
    
    const dropdownRow = sheet.getRange(2, 1, 1, targetLastColumn);
    if (!dropdownRow.isPartOfMerge()) {
      dropdownRow.merge();
    }
    
    const attendanceHeaderRow = sheet.getRange(3, 1, 1, targetLastColumn);
    if (!attendanceHeaderRow.isPartOfMerge()) {
      attendanceHeaderRow.merge();
    }
    
    // Log success
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Successfully expanded sheet for ${studentCount} students`, 'INFO', 'expandSheetForStudents');
    }
  } catch (error) {
    // Log the error but don't throw - we'll still try to proceed with the sheet as is
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error expanding sheet: ${error.message}`, 'ERROR', 'expandSheetForStudents');
    } else {
      Logger.log(`Error expanding sheet: ${error.message}`);
    }
  }
}

/**
 * Populates class dates in the Group Lesson Tracker sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param daxkoData - All data from the Daxko sheet
 * @param matchingRow - The row index matching the selected class
 */
function populateClassDates(sheet, daxkoData, matchingRow) {
  try {
    // Get start date (AB column) and end date (AC column)
    const startDateValue = daxkoData[matchingRow][27]; // Column AB
    const endDateValue = daxkoData[matchingRow][28];   // Column AC
    
    // Convert to JavaScript Dates
    let startDate = null;
    let endDate = null;
    
    if (startDateValue instanceof Date) {
      startDate = startDateValue;
    } else if (typeof startDateValue === 'string') {
      startDate = new Date(startDateValue);
    } else if (typeof startDateValue === 'number') {
      // Handle Excel serial date
      startDate = new Date(Math.round((startDateValue - 25569) * 86400 * 1000));
    }
    
    if (endDateValue instanceof Date) {
      endDate = endDateValue;
    } else if (typeof endDateValue === 'string') {
      endDate = new Date(endDateValue);
    } else if (typeof endDateValue === 'number') {
      // Handle Excel serial date
      endDate = new Date(Math.round((endDateValue - 25569) * 86400 * 1000));
    }
    
    // Check if dates are valid
    if (!startDate || isNaN(startDate.getTime()) || !endDate || isNaN(endDate.getTime())) {
      Logger.log('Invalid start or end date');
      return;
    }
    
    // Get day of week (0 = Sunday, 1 = Monday, etc.)
    const dayOfWeek = startDate.getDay();
    
    // Generate weekly dates from start to end
    const classDates = [];
    const currentDate = new Date(startDate);
    
    // Add first date
    classDates.push(new Date(currentDate));
    
    // Generate weekly dates
    while (currentDate < endDate) {
      // Add 7 days for next class
      currentDate.setDate(currentDate.getDate() + 7);
      
      // If we haven't gone past the end date, add this date
      if (currentDate <= endDate) {
        classDates.push(new Date(currentDate));
      }
    }
    
    // Format and populate class dates
    const dateFormatter = new Intl.DateTimeFormat('en-US', {
      month: 'short',
      day: 'numeric',
      year: 'numeric'
    });
    
    // Clear existing dates
    for (let i = 5; i <= 12; i++) {
      sheet.getRange(`A${i}`).setValue('');
    }
    
    // Populate dates (up to 8 classes)
    const maxDates = Math.min(classDates.length, 8);
    for (let i = 0; i < maxDates; i++) {
      const formattedDate = dateFormatter.format(classDates[i]);
      sheet.getRange(`A${i + 5}`).setValue(`Class ${i + 1}: ${formattedDate}`);
    }
    
    // Log the number of dates generated
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Populated ${maxDates} class dates in Group Lesson Tracker`, 'INFO', 'populateClassDates');
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error populating class dates: ${error.message}`, 'ERROR', 'populateClassDates');
    } else {
      Logger.log(`Error populating class dates: ${error.message}`);
    }
  }
}

/**
 * Populates class skills in the Group Lesson Tracker sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param programAbbreviation - The program abbreviation (e.g., "S1")
 */
/**
 * Populates class skills in the Group Lesson Tracker sheet
 * and returns mapping of skills to their column indices
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param programAbbreviation - The program abbreviation (e.g., "S1")
 * @returns Object mapping skill rows in the Group Lesson Tracker to column indices in SwimmerSkills
 */
function populateClassSkills(sheet, programAbbreviation) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swimmerSkillsSheet = ss.getSheetByName('SwimmerSkills');
    
    if (!swimmerSkillsSheet) {
      Logger.log('SwimmerSkills sheet not found');
      return null;
    }
    
    // Get all skills from SwimmerSkills sheet
    const skillsData = swimmerSkillsSheet.getDataRange().getValues();
    
    // Extract header row to find skills that start with program abbreviation
    const headers = skillsData[0];
    const matchingSkills = [];
    
    // Create a mapping of row numbers to skill columns
    // This will be used later to populate student skills
    const skillsMapping = {};
    
    for (let i = 0; i < headers.length; i++) {
      const headerText = headers[i] ? headers[i].toString() : '';
      if (headerText.startsWith(programAbbreviation)) {
        matchingSkills.push({
          index: i,
          name: headerText,
          column: i
        });
      }
    }
    
    // Clear existing skills in first section (A14-A30)
    for (let i = 14; i <= 30; i++) {
      sheet.getRange(`A${i}`).setValue('');
    }
    
    // Populate skills in first section (up to 17 skills now with the extra row)
    const maxSkillsFirstSection = Math.min(matchingSkills.length, 17);
    for (let i = 0; i < maxSkillsFirstSection; i++) {
      const rowNum = i + 14;
      sheet.getRange(`A${rowNum}`).setValue(matchingSkills[i].name)
        .setFontWeight('bold')
        .setBackground('#D9EAD3');
      
      // Store the mapping of row number to skill column in SwimmerSkills sheet
      skillsMapping[rowNum] = matchingSkills[i].column;
    }
    
    // If there are more skills, continue in second section (A32-A41)
    if (matchingSkills.length > 17) {
      // Clear existing skills in second section
      for (let i = 32; i <= 41; i++) {
        sheet.getRange(`A${i}`).setValue('');
      }
      
      // Calculate remaining skills to display
      const remainingSkills = Math.min(matchingSkills.length - 17, 10);
      
      for (let i = 0; i < remainingSkills; i++) {
        const skillIndex = i + 17; // Start from the 18th skill
        const rowNum = i + 32; // Start from row 32
        
        sheet.getRange(`A${rowNum}`).setValue(matchingSkills[skillIndex].name)
          .setFontWeight('bold')
          .setBackground('#D9EAD3');
        
        // Store the mapping of row number to skill column in SwimmerSkills sheet
        skillsMapping[rowNum] = matchingSkills[skillIndex].column;
      }
    }
    
    // Log the number of skills found and populated
    const totalSkillsPopulated = Math.min(matchingSkills.length, 27); // 17 in first section + 10 in second
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Populated ${totalSkillsPopulated} skills for program ${programAbbreviation} in Group Lesson Tracker`, 'INFO', 'populateClassSkills');
    }
    
    // Return the mapping of Group Lesson Tracker rows to SwimmerSkills columns
    return {
      skillsMapping: skillsMapping,
      swimmerSkillsData: skillsData
    };
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error populating class skills: ${error.message}`, 'ERROR', 'populateClassSkills');
    } else {
      Logger.log(`Error populating class skills: ${error.message}`);
    }
    return null;
  }
}

/**
 * Populates student skills from the SwimmerSkills sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param studentsData - Array of student data objects with names and columns
 * @param skillsData - Object containing skills mapping and SwimmerSkills data
 */
function populateStudentSkills(sheet, studentsData, skillsData) {
  try {
    if (!studentsData || studentsData.length === 0 || !skillsData || !skillsData.skillsMapping) {
      Logger.log('Missing data for populating student skills');
      return;
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swimmerSkillsData = skillsData.swimmerSkillsData;
    const skillsMapping = skillsData.skillsMapping;
    
    Logger.log(`Starting to populate skills for ${studentsData.length} students`);
    Logger.log(`SwimmerSkills data has ${swimmerSkillsData.length} rows`);
    
    // For each student
    for (let i = 0; i < studentsData.length; i++) {
      const student = studentsData[i];
      
      // Only proceed if we have firstName and beginningColumn
      if (!student.firstName || !student.beginningColumn) {
        Logger.log(`Skipping student with missing data: ${JSON.stringify(student)}`);
        continue;
      }
      
      Logger.log(`Finding student: ${student.firstName} ${student.lastName}`);
      
      // First try to find the student in the SwimmerSkills sheet
      let studentRow = -1;
      
      // Try to match first name and last name
      for (let row = 1; row < swimmerSkillsData.length; row++) {
        const firstName = swimmerSkillsData[row][0]; // Column A
        const lastName = swimmerSkillsData[row][1];  // Column B
        
        // Skip empty rows
        if (!firstName && !lastName) continue;
        
        // Convert values to strings and trim whitespace for comparison
        const ssFirstName = firstName ? firstName.toString().trim().toLowerCase() : '';
        const ssLastName = lastName ? lastName.toString().trim().toLowerCase() : '';
        const studentFirstName = student.firstName.toString().trim().toLowerCase();
        const studentLastName = student.lastName ? student.lastName.toString().trim().toLowerCase() : '';
        
        // Check for exact match on first and last name
        if (ssFirstName === studentFirstName && ssLastName === studentLastName) {
          studentRow = row;
          Logger.log(`Found exact match for ${studentFirstName} ${studentLastName} at row ${row+1}`);
          break;
        }
      }
      
      // If no match was found, try again with only first name match
      if (studentRow === -1 && student.firstName) {
        for (let row = 1; row < swimmerSkillsData.length; row++) {
          const firstName = swimmerSkillsData[row][0]; // Column A
          
          // Skip empty rows
          if (!firstName) continue;
          
          // Convert values to strings and trim whitespace for comparison
          const ssFirstName = firstName ? firstName.toString().trim().toLowerCase() : '';
          const studentFirstName = student.firstName.toString().trim().toLowerCase();
          
          // Check for match on first name only if we couldn't find exact match
          if (ssFirstName === studentFirstName) {
            studentRow = row;
            Logger.log(`Found first-name-only match for ${studentFirstName} at row ${row+1}`);
            break;
          }
        }
      }
      
      // If we found the student, populate their skills
      if (studentRow !== -1) {
        Logger.log(`Populating skills for ${student.firstName} ${student.lastName} from row ${studentRow+1} in SwimmerSkills`);
        
        // Loop through each skill in the mapping
        Object.keys(skillsMapping).forEach(rowNum => {
          const skillColIndex = skillsMapping[rowNum]; // Column index in SwimmerSkills
          const skillName = swimmerSkillsData[0][skillColIndex]; // Get skill name from header row
          
          // Get the skill value for this student
          const skillValue = swimmerSkillsData[studentRow][skillColIndex];
          
          // If there's a value, populate it in the beginning column for this student
          if (skillValue !== null && skillValue !== undefined) {
            const beginningCell = student.beginningColumn + rowNum;
            
            try {
              sheet.getRange(beginningCell).setValue(skillValue);
              Logger.log(`Set skill "${skillName}" to value "${skillValue}" for ${student.firstName} at cell ${beginningCell}`);
              
              // Apply color based on value
              if (skillValue === 'X') {
                sheet.getRange(beginningCell).setBackground('#D9EAD3'); // Light green for performed the skill
              } else if (skillValue === '/') {
                sheet.getRange(beginningCell).setBackground('#FFF2CC'); // Light yellow for taught the skill
              } else {
                sheet.getRange(beginningCell).setBackground(null); // Clear background for blank
              }
            } catch (e) {
              Logger.log(`Error setting skill value in cell ${beginningCell}: ${e.message}`);
            }
          }
        });
      } else {
        Logger.log(`Could not find student ${student.firstName} ${student.lastName} in SwimmerSkills sheet`);
      }
    }
    
    // Log success
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Populated skills for ${studentsData.length} students in Group Lesson Tracker`, 'INFO', 'populateStudentSkills');
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error populating student skills: ${error.message}`, 'ERROR', 'populateStudentSkills');
    } else {
      Logger.log(`Error populating student skills: ${error.message}`);
    }
  }
}

/**
 * Synchronizes student data between Group Lesson Tracker and SwimmerSkills
 * Works bidirectionally:
 * 1. Imports skills from SwimmerSkills to Group Lesson Tracker
 * 2. Exports updated skills from Group Lesson Tracker to SwimmerSkills
 * 
 * @param sheet - The Group Lesson Tracker sheet
 */
function syncStudentDataWithSwimmerSkills(sheet) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swimmerSkillsSheet = ss.getSheetByName('SwimmerSkills');
    
    if (!swimmerSkillsSheet) {
      Logger.log('SwimmerSkills sheet not found');
      SpreadsheetApp.getUi().alert('Error', 'SwimmerSkills sheet not found. Unable to sync data.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get all data from SwimmerSkills sheet
    const swimmerSkillsData = swimmerSkillsSheet.getDataRange().getValues();
    
    // First, collect all the students and skills from the Group Lesson Tracker
    const students = collectStudentsFromGroupLessonTracker(sheet);
    
    if (!students || students.length === 0) {
      Logger.log('No students found in Group Lesson Tracker');
      SpreadsheetApp.getUi().alert('No Data to Sync', 'No student data found to sync. Please make sure students are populated in the sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Get the skills mapping from Group Lesson Tracker
    const skills = collectSkillsFromGroupLessonTracker(sheet);
    
    if (!skills || Object.keys(skills).length === 0) {
      Logger.log('No skills found in Group Lesson Tracker');
      SpreadsheetApp.getUi().alert('No Skills to Sync', 'No skills found to sync. Please make sure skills are populated in the sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Now, for each student, find their corresponding row in SwimmerSkills
    let syncCount = 0;
    let insertCount = 0;
    let updateCount = 0;
    
    // Track students that couldn't be found
    const notFoundStudents = [];
    
    for (const student of students) {
      // Find student in SwimmerSkills sheet
      let studentRow = -1;
      
      // Try to match first name and last name
      for (let row = 1; row < swimmerSkillsData.length; row++) {
        const firstName = swimmerSkillsData[row][0]; // Column A
        const lastName = swimmerSkillsData[row][1];  // Column B
        
        // Skip empty rows
        if (!firstName && !lastName) continue;
        
        // Check if this is our student (case-insensitive match)
        if (firstName && lastName && 
            firstName.toString().trim().toLowerCase() === student.firstName.toString().trim().toLowerCase() && 
            lastName.toString().trim().toLowerCase() === student.lastName.toString().trim().toLowerCase()) {
          studentRow = row;
          break;
        }
      }
      
      // If student not found, we'll add them to the not found list
      if (studentRow === -1) {
        notFoundStudents.push(student.fullName);
        continue;
      }
      
      // For each skill, check the student's end column value in Group Lesson Tracker
      // and update the corresponding cell in SwimmerSkills
      for (const [rowNum, skill] of Object.entries(skills)) {
        // Get the value from the student's "ending" column (for completed assessment)
        try {
          // Calculate the cell reference for the ending column (e.g., "C14")
          const endColumn = student.endingColumn || '';
          if (!endColumn) continue;
          
          const endCellRef = endColumn + rowNum;
          const endValue = sheet.getRange(endCellRef).getValue();
          
          // If there's a value in the ending column, update it in SwimmerSkills
          if (endValue && endValue.toString().trim() !== '') {
            // The column in SwimmerSkills is skill.column
            const skillColumn = skill.column;
            
            // Check if there's already a value in this cell
            const currentValue = swimmerSkillsData[studentRow][skillColumn];
            
            // Update the value in SwimmerSkills sheet
            swimmerSkillsSheet.getRange(studentRow + 1, skillColumn + 1).setValue(endValue);
            
            if (currentValue && currentValue.toString().trim() !== '') {
              updateCount++;
            } else {
              insertCount++;
            }
            
            syncCount++;
          }
        } catch (e) {
          Logger.log(`Error syncing skill for student ${student.fullName}, row ${rowNum}: ${e.message}`);
        }
      }
    }
    
    // Now pull in any new values from SwimmerSkills to the beginning columns
    pullDataFromSwimmerSkills(sheet, students, skills, swimmerSkillsData);
    
    // Show result to user
    if (notFoundStudents.length > 0) {
      const message = `Synced ${syncCount} skills (${insertCount} new, ${updateCount} updated).\n\n` +
                     `Legend:\n- 'X' = Student performed skill (green)\n- '/' = Skill was taught (yellow)\n\n` +
                     `${notFoundStudents.length} student(s) not found in SwimmerSkills:\n` +
                     notFoundStudents.join(', ');
      
      SpreadsheetApp.getUi().alert('Sync Completed with Warnings', message, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      const message = `Successfully synced ${syncCount} skills (${insertCount} new, ${updateCount} updated) with SwimmerSkills.\n\n` +
                      `Legend:\n- 'X' = Student performed skill (green)\n- '/' = Skill was taught (yellow)`;
      SpreadsheetApp.getUi().alert('Sync Completed', message, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
    // Log the sync
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Synced ${syncCount} skills with SwimmerSkills (${insertCount} new, ${updateCount} updated)`, 'INFO', 'syncStudentDataWithSwimmerSkills');
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error syncing with SwimmerSkills: ${error.message}`, 'ERROR', 'syncStudentDataWithSwimmerSkills');
    } else {
      Logger.log(`Error syncing with SwimmerSkills: ${error.message}`);
    }
    SpreadsheetApp.getUi().alert('Sync Error', `An error occurred while syncing with SwimmerSkills: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Collects student data from the Group Lesson Tracker sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @returns Array of student data objects
 */
function collectStudentsFromGroupLessonTracker(sheet) {
  // Define the columns for student names (B, D, F, etc.)
  const studentNameColumns = ['B', 'D', 'F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'V', 'X', 
                             'Z', 'AB', 'AD', 'AF', 'AH', 'AJ', 'AL', 'AN', 'AP', 'AR', 'AT', 'AV'];
  
  // The ending columns (C, E, G, etc.)
  const studentEndColumns = ['C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y',
                            'AA', 'AC', 'AE', 'AG', 'AI', 'AK', 'AM', 'AO', 'AQ', 'AS', 'AU', 'AW'];
  
  // Collect all the students
  const students = [];
  
  // Log for debugging
  Logger.log('Collecting students from Group Lesson Tracker');
  
  // Loop through potential student positions
  for (let i = 0; i < studentNameColumns.length; i++) {
    try {
      // Get student name from the header cell
      const studentName = sheet.getRange(studentNameColumns[i] + '4').getValue();
      
      // If this position has a name, add it
      if (studentName && studentName.toString().trim() !== '') {
        // Parse first and last name
        const nameParts = studentName.toString().trim().split(' ');
        let firstName = '';
        let lastName = '';
        
        if (nameParts.length === 1) {
          firstName = nameParts[0];
          lastName = ''; // Ensure lastName is set even if empty
        } else if (nameParts.length >= 2) {
          firstName = nameParts[0];
          lastName = nameParts.slice(1).join(' ');
        }
        
        // Create student object with all needed data
        const student = {
          fullName: studentName.toString().trim(),
          firstName: firstName,
          lastName: lastName,
          beginningColumn: studentNameColumns[i],
          endingColumn: studentEndColumns[i]
        };
        
        students.push(student);
        Logger.log(`Found student: ${firstName} ${lastName} in column ${studentNameColumns[i]}`);
      }
    } catch (e) {
      // Cell might not exist yet, just continue
      Logger.log(`Error reading student in column ${studentNameColumns[i]}: ${e.message}`);
      continue;
    }
  }
  
  Logger.log(`Collected ${students.length} students total`);
  return students;
}

/**
 * Collects skills data from the Group Lesson Tracker sheet
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @returns Object mapping row numbers to skill names and columns
 */
function collectSkillsFromGroupLessonTracker(sheet) {
  const skills = {};
  Logger.log('Starting to collect skills from Group Lesson Tracker');
  
  // Examine skill rows (A14-A43) - expanded from A14-A29 to capture more skills
  for (let row = 14; row <= 43; row++) {
    try {
      const skillName = sheet.getRange(`A${row}`).getValue();
      
      if (skillName && skillName.toString().trim() !== '') {
        // We need to find the corresponding column in SwimmerSkills
        const skill = {
          name: skillName.toString().trim(),
          row: row
        };
        
        Logger.log(`Found skill "${skill.name}" in row ${row}`);
        
        // Determine the column in SwimmerSkills sheet
        // We'll do this by matching the skill name to headers
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const swimmerSkillsSheet = ss.getSheetByName('SwimmerSkills');
        
        if (swimmerSkillsSheet) {
          // Get the header row from SwimmerSkills
          const headers = swimmerSkillsSheet.getRange(1, 1, 1, swimmerSkillsSheet.getLastColumn()).getValues()[0];
          
          // Find the matching column (exact match first)
          for (let i = 0; i < headers.length; i++) {
            if (headers[i] && headers[i].toString().trim() === skillName.toString().trim()) {
              skill.column = i;
              Logger.log(`Found exact match for skill "${skill.name}" in SwimmerSkills column ${i+1}`);
              break;
            }
          }
          
          // If we didn't find an exact match, try substring matching
          if (skill.column === undefined) {
            // First try finding skills where header contains the Group Lesson Tracker skill name
            for (let i = 0; i < headers.length; i++) {
              if (headers[i] && headers[i].toString().trim().toLowerCase().includes(skillName.toString().trim().toLowerCase())) {
                skill.column = i;
                Logger.log(`Found partial match for skill "${skill.name}" in SwimmerSkills column ${i+1} (${headers[i]})`);
                break;
              }
            }
            
            // If still not found, try the reverse - Group Lesson Tracker skill name contains header
            if (skill.column === undefined) {
              for (let i = 0; i < headers.length; i++) {
                if (headers[i] && skillName.toString().trim().toLowerCase().includes(headers[i].toString().trim().toLowerCase())) {
                  skill.column = i;
                  Logger.log(`Found reverse partial match for skill "${skill.name}" in SwimmerSkills column ${i+1} (${headers[i]})`);
                  break;
                }
              }
            }
          }
          
          // If we found a column, add this skill to the map
          if (skill.column !== undefined) {
            skills[row] = skill;
          } else {
            Logger.log(`Could not find matching column in SwimmerSkills for skill "${skill.name}"`);
          }
        }
      }
    } catch (e) {
      // Just log the error and continue
      Logger.log(`Error processing skill in row ${row}: ${e.message}`);
      continue;
    }
  }
  
  Logger.log(`Collected ${Object.keys(skills).length} skills with matched columns`);
  return skills;
}

/**
 * Pulls updated data from SwimmerSkills into the Group Lesson Tracker's beginning columns
 * 
 * @param sheet - The Group Lesson Tracker sheet
 * @param students - Array of student data objects
 * @param skills - Object mapping row numbers to skill data
 * @param swimmerSkillsData - 2D array of data from SwimmerSkills sheet
 */
function pullDataFromSwimmerSkills(sheet, students, skills, swimmerSkillsData) {
  Logger.log('Starting to pull data from SwimmerSkills to Group Lesson Tracker');
  
  // For each student
  for (const student of students) {
    // Find student in SwimmerSkills sheet
    let studentRow = -1;
    
    Logger.log(`Looking up student ${student.firstName} ${student.lastName} in SwimmerSkills`);
    
    // Try to match first name and last name
    for (let row = 1; row < swimmerSkillsData.length; row++) {
      const firstName = swimmerSkillsData[row][0]; // Column A
      const lastName = swimmerSkillsData[row][1];  // Column B
      
      // Skip empty rows
      if (!firstName && !lastName) continue;
      
      // Convert values to strings and trim whitespace for comparison
      const ssFirstName = firstName ? firstName.toString().trim().toLowerCase() : '';
      const ssLastName = lastName ? lastName.toString().trim().toLowerCase() : '';
      const studentFirstName = student.firstName ? student.firstName.toString().trim().toLowerCase() : '';
      const studentLastName = student.lastName ? student.lastName.toString().trim().toLowerCase() : '';
      
      // Check for exact match on first and last name
      if (ssFirstName === studentFirstName && ssLastName === studentLastName) {
        studentRow = row;
        Logger.log(`Found exact match for ${studentFirstName} ${studentLastName} at row ${row+1}`);
        break;
      }
    }
    
    // If no match was found, try again with only first name match
    if (studentRow === -1 && student.firstName) {
      for (let row = 1; row < swimmerSkillsData.length; row++) {
        const firstName = swimmerSkillsData[row][0]; // Column A
        
        // Skip empty rows
        if (!firstName) continue;
        
        // Convert values to strings and trim whitespace for comparison
        const ssFirstName = firstName ? firstName.toString().trim().toLowerCase() : '';
        const studentFirstName = student.firstName.toString().trim().toLowerCase();
        
        // Check for match on first name only if we couldn't find exact match
        if (ssFirstName === studentFirstName) {
          studentRow = row;
          Logger.log(`Found first-name-only match for ${studentFirstName} at row ${row+1}`);
          break;
        }
      }
    }
    
    // If student not found, skip
    if (studentRow === -1) {
      Logger.log(`Could not find student ${student.firstName} ${student.lastName} in SwimmerSkills sheet`);
      continue;
    }
    
    // For each skill, get the value from SwimmerSkills and update the beginning column
    for (const [rowNum, skill] of Object.entries(skills)) {
      try {
        // Get the value from SwimmerSkills
        const skillValue = swimmerSkillsData[studentRow][skill.column];
        const skillName = swimmerSkillsData[0][skill.column]; // Get skill name from header row
        
        // If there's a value, update the beginning column
        if (skillValue !== null && skillValue !== undefined) {
          const beginningCell = student.beginningColumn + rowNum;
          
          // Get the current value in the beginning column
          const currentValue = sheet.getRange(beginningCell).getValue();
          
          // Only update if different (to avoid unnecessary updates)
          if (currentValue !== skillValue) {
            sheet.getRange(beginningCell).setValue(skillValue);
            Logger.log(`Updated ${student.firstName}'s skill "${skillName}" to "${skillValue}" in cell ${beginningCell}`);
            
            // Apply color based on value
            if (skillValue === 'X') {
              sheet.getRange(beginningCell).setBackground('#D9EAD3'); // Light green for performed the skill
            } else if (skillValue === '/') {
              sheet.getRange(beginningCell).setBackground('#FFF2CC'); // Light yellow for taught the skill
            } else {
              sheet.getRange(beginningCell).setBackground(null); // Clear background for blank
            }
          }
        }
      } catch (e) {
        Logger.log(`Error pulling skill for student ${student.fullName}, row ${rowNum}: ${e.message}`);
      }
    }
  }
  Logger.log('Finished pulling data from SwimmerSkills to Group Lesson Tracker');
}

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
  safeSetProperty: safeSetProperty,
  handleClassesSheetEdit: handleClassesSheetEdit,
  handleAnnouncementsSheetEdit: handleAnnouncementsSheetEdit,
  handleGroupLessonTrackerDropdownChange: handleGroupLessonTrackerDropdownChange,
  populateStudentNames: populateStudentNames,
  populateClassDates: populateClassDates,
  populateClassSkills: populateClassSkills,
  populateStudentSkills: populateStudentSkills,
  syncStudentDataWithSwimmerSkills: syncStudentDataWithSwimmerSkills,
  collectStudentsFromGroupLessonTracker: collectStudentsFromGroupLessonTracker,
  collectSkillsFromGroupLessonTracker: collectSkillsFromGroupLessonTracker,
  pullDataFromSwimmerSkills: pullDataFromSwimmerSkills,
  getProgramAbbreviation: getProgramAbbreviation,
  expandSheetForStudents: expandSheetForStudents,
  truncateSheetToOriginalSize: truncateSheetToOriginalSize,
  truncateSheetToSize: truncateSheetToSize,
  clearExistingTrackerData: clearExistingTrackerData,
  applyAlternatingBackground: applyAlternatingBackground,
  columnToLetter: columnToLetter
};