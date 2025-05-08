/**
 * YSL Hub v2 Data Integration Module
 * 
 * This module handles all data flow operations between system components, including roster
 * extraction, assessment criteria synchronization, and bi-directional integration with
 * external systems such as the YSL Swimmer Log workbook.
 * 
 * @author PenBay YMCA
 * @version 2.0
 * @date 2025-04-27
 */

/**
 * Initialize data structures based on configuration parameters.
 * Creates necessary sheets and imports initial data from external sources.
 * 
 * @param {Object} config - Configuration parameters
 * @return {boolean} Success status
 */
function initializeDataStructures(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Log initialization start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting data structures initialization', 'INFO', 'initializeDataStructures');
    }
    
    // Create or clear Daxko sheet
    createOrClearSheet(ss, 'Daxko');
    
    // Create or clear Classes sheet
    createOrClearSheet(ss, 'Classes');
    
    // Import roster data from the session roster file
    importSessionRosterData(config.sessionName, config.rosterFolderUrl);
    
    // Import assessment criteria from Swimmer Log
    importAssessmentCriteria(config.swimmerRecordsUrl);
    
    // Set up class selector in Classes sheet
    setupClassSelector();
    
    // Import instructor data from Session Programs workbook if available
    try {
      if (config.sessionProgramsUrl) {
        importInstructorData(config.sessionProgramsUrl);
      }
    } catch (error) {
      // Log but continue - instructor data is helpful but not critical
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Optional instructor data import error: ${error.message}`, 'WARNING', 'initializeDataStructures');
      } else {
        Logger.log(`Optional instructor data import error: ${error.message}`);
      }
    }
    
    // Log successful initialization
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Data structures initialization completed successfully', 'INFO', 'initializeDataStructures');
    }
    
    return true;
  } catch (error) {
    // Handle and log initialization errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'initializeDataStructures', 
        'Failed to initialize data structures. Please check your configuration and try again.');
    } else {
      Logger.log(`Data structure initialization error: ${error.message}`);
      throw new Error(`Failed to initialize data structures: ${error.message}`);
    }
    return false;
  }
}

/**
 * Creates a new sheet or clears an existing one.
 * 
 * @param {Spreadsheet} ss - The active spreadsheet
 * @param {string} sheetName - Name of the sheet to create or clear
 * @return {Sheet} The created or cleared sheet
 */
function createOrClearSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    // Clear existing sheet
    sheet.clear();
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Cleared existing sheet: ${sheetName}`, 'INFO', 'createOrClearSheet');
    }
  } else {
    // Create new sheet
    sheet = ss.insertSheet(sheetName);
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Created new sheet: ${sheetName}`, 'INFO', 'createOrClearSheet');
    }
  }
  
  return sheet;
}

/**
 * Imports data from the session roster file in the specified folder.
 * Enhanced with better error handling and logging.
 * 
 * @param {string} sessionName - The session name
 * @param {string} folderUrl - URL of the folder containing session rosters
 * @return {boolean} Success status
 */
function importSessionRosterData(sessionName, folderUrl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const daxkoSheet = ss.getSheetByName('Daxko');
  
  if (!daxkoSheet) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Daxko sheet not found', 'ERROR', 'importSessionRosterData');
    }
    throw new Error('Daxko sheet not found');
  }
  
  try {
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Importing roster data for session: ${sessionName}`, 'INFO', 'importSessionRosterData');
    }
    
    // Extract folder ID from URL with improved function
    let folderId;
    if (typeof GlobalFunctions.extractIdFromUrl === 'function') {
      folderId = GlobalFunctions.extractIdFromUrl(folderUrl);
    } else {
      // Fallback to basic extraction
      folderId = extractIdFromUrl(folderUrl);
    }
    
    if (!folderId) {
      // Try using the default folder ID as a fallback
      folderId = '1vlR8WwEyLWOuO-JUzCzrdLikv-hVVld4';
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Invalid folder URL, using default folder ID: ${folderId}`, 'WARNING', 'importSessionRosterData');
      }
    } else {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Using folder ID: ${folderId}`, 'INFO', 'importSessionRosterData');
      }
    }
    
    // Get the folder
    let folder;
    if (typeof GlobalFunctions.safeGetFolderById === 'function') {
      folder = GlobalFunctions.safeGetFolderById(folderId);
    } else {
      // Fallback to direct access
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (error) {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Failed to access folder: ${error.message}`, 'ERROR', 'importSessionRosterData');
        }
        throw new Error(`Folder access error: ${error.message}`);
      }
    }
    
    if (!folder) {
      throw new Error('Folder not found or access denied');
    }
    
    // Construct the expected roster filename pattern
    const rosterFilePattern = `YSL ${sessionName} Roster`;
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Looking for roster file matching pattern: ${rosterFilePattern}`, 'INFO', 'importSessionRosterData');
    }
    
    // Find the roster file
    let rosterFile = null;
    const files = folder.getFiles();
    let fileCount = 0;
    let filesList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      fileCount++;
      filesList.push(file.getName());
      
      if (file.getName().includes(rosterFilePattern)) {
        rosterFile = file;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Found roster file: ${file.getName()}`, 'INFO', 'importSessionRosterData');
        }
        
        break;
      }
    }
    
    if (!rosterFile) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`No roster file found matching pattern "${rosterFilePattern}" in folder with ID ${folderId}`, 'ERROR', 'importSessionRosterData');
        ErrorHandling.logMessage(`Files in folder: ${filesList.join(", ")}`, 'DEBUG', 'importSessionRosterData');
      }
      
      throw new Error(`Roster file matching pattern "${rosterFilePattern}" not found in the specified folder`);
    }
    
    // Check if the file is a spreadsheet or CSV
    const fileType = rosterFile.getMimeType();
    let csvData;
    
    if (fileType === 'application/vnd.google-apps.spreadsheet') {
      // It's a Google Sheet - open and get data directly
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Processing roster as Google Sheet', 'INFO', 'importSessionRosterData');
      }
      
      const rosterSS = SpreadsheetApp.openById(rosterFile.getId());
      const rosterSheet = rosterSS.getSheets()[0]; // Use the first sheet
      csvData = rosterSheet.getDataRange().getValues();
    } else {
      // Assume it's a CSV or other text format
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Processing roster as CSV file', 'INFO', 'importSessionRosterData');
      }
      
      // Read the roster file content
      const content = rosterFile.getBlob().getDataAsString();
      
      // Parse CSV content
      csvData = Utilities.parseCsv(content);
    }
    
    // Check if we got any data
    if (!csvData || csvData.length === 0) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Roster file appears to be empty', 'ERROR', 'importSessionRosterData');
      }
      throw new Error('Roster file appears to be empty');
    }
    
    // Count columns and rows for logging
    const rowCount = csvData.length;
    const colCount = csvData[0].length;
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Processing roster data with ${rowCount} rows and ${colCount} columns`, 'INFO', 'importSessionRosterData');
    }
    
    // Write data to Daxko sheet
    daxkoSheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
    
    // Format header row
    daxkoSheet.getRange(1, 1, 1, csvData[0].length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Freeze header row
    daxkoSheet.setFrozenRows(1);
    
    // Auto-resize columns for better visibility
    for (let i = 1; i <= csvData[0].length; i++) {
      daxkoSheet.autoResizeColumn(i);
    }
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Roster data imported successfully', 'INFO', 'importSessionRosterData');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'importSessionRosterData', 
        'Failed to import roster data. Please check file access and try again.');
    } else {
      Logger.log(`Roster data import error: ${error.message}`);
      throw error; // Re-throw to caller
    }
    return false;
  }
}

/**
 * Imports assessment criteria from the YSL Swimmer Log workbook.
 * Enhanced with better error handling and logging.
 * 
 * @param {string} swimmerLogUrl - URL of the Swimmer Log workbook
 * @return {boolean} Success status
 */
function importAssessmentCriteria(swimmerLogUrl) {
  try {
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Importing assessment criteria', 'INFO', 'importAssessmentCriteria');
    }
    
    // Extract spreadsheet ID from URL
    let ssId;
    if (typeof GlobalFunctions.extractIdFromUrl === 'function') {
      ssId = GlobalFunctions.extractIdFromUrl(swimmerLogUrl);
    } else {
      // Fallback to basic extraction
      ssId = extractIdFromUrl(swimmerLogUrl);
    }
    
    if (!ssId) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Invalid Swimmer Log URL', 'ERROR', 'importAssessmentCriteria');
      }
      throw new Error('Invalid Swimmer Log URL');
    }
    
    // Open the Swimmer Log workbook
    let swimmerLog;
    if (typeof GlobalFunctions.safeGetSpreadsheetById === 'function') {
      swimmerLog = GlobalFunctions.safeGetSpreadsheetById(ssId);
    } else {
      // Fallback to direct access
      swimmerLog = SpreadsheetApp.openById(ssId);
    }
    
    if (!swimmerLog) {
      throw new Error('Could not open Swimmer Log workbook. Please check the URL and permissions.');
    }
    
    // Get the Swimmers sheet
    const swimmersSheet = swimmerLog.getSheetByName('Swimmers');
    
    if (!swimmersSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Swimmers sheet not found in the Swimmer Log workbook', 'ERROR', 'importAssessmentCriteria');
      }
      throw new Error('Swimmers sheet not found in the Swimmer Log workbook');
    }
    
    // Get assessment criteria from row 2
    const criteriaRow = swimmersSheet.getRange(2, 1, 1, swimmersSheet.getLastColumn()).getValues()[0];
    const criteriaHeaders = swimmersSheet.getRange(1, 1, 1, swimmersSheet.getLastColumn()).getValues()[0];
    
    // Store criteria for later use in class sheets
    const scriptProperties = PropertiesService.getScriptProperties();
    scriptProperties.setProperty('assessmentCriteria', JSON.stringify(criteriaRow));
    scriptProperties.setProperty('criteriaHeaders', JSON.stringify(criteriaHeaders));
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Assessment criteria imported successfully: ${criteriaHeaders.length} columns`, 'INFO', 'importAssessmentCriteria');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'importAssessmentCriteria', 
        'Failed to import assessment criteria. Please check access to the Swimmer Log and try again.');
    } else {
      Logger.log(`Assessment criteria import error: ${error.message}`);
      throw new Error(`Failed to import assessment criteria: ${error.message}`);
    }
    return false;
  }
}

/**
 * Sets up the class selector in the Classes sheet.
 * Enhanced with better error handling and logging.
 * 
 * @return {boolean} Success status
 */
function setupClassSelector() {
  try {
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Setting up class selector', 'INFO', 'setupClassSelector');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    const daxkoSheet = ss.getSheetByName('Daxko');
    
    if (!classesSheet || !daxkoSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Required sheets not found', 'ERROR', 'setupClassSelector');
      }
      throw new Error('Required sheets not found');
    }
    
    // Find the column indices for relevant fields in Daxko sheet
    const daxkoHeaders = daxkoSheet.getRange(1, 1, 1, daxkoSheet.getLastColumn()).getValues()[0];
    
    // Use the findColumnIndex function if available, otherwise use indexOf
    let programColIndex, dayColIndex, timeColIndex, siteColIndex;
    
    if (typeof GlobalFunctions.findColumnIndex === 'function') {
      programColIndex = GlobalFunctions.findColumnIndex(daxkoHeaders, ['program', 'class', 'stage', 'level']);
      dayColIndex = GlobalFunctions.findColumnIndex(daxkoHeaders, ['day', 'days', 'day(s)', 'weekday']);
      timeColIndex = GlobalFunctions.findColumnIndex(daxkoHeaders, ['time', 'session time', 'class time']);
      siteColIndex = GlobalFunctions.findColumnIndex(daxkoHeaders, ['location', 'site', 'facility']);
    } else {
      // Fallback to basic indexOf
      programColIndex = daxkoHeaders.indexOf('Program');
      dayColIndex = daxkoHeaders.indexOf('Day(s) of Week');
      timeColIndex = daxkoHeaders.indexOf('Session Time');
      siteColIndex = daxkoHeaders.indexOf('Site');
    }
    
    if (programColIndex === -1 || dayColIndex === -1 || timeColIndex === -1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Required columns not found in Daxko sheet. Headers: ${daxkoHeaders.join(', ')}`, 'ERROR', 'setupClassSelector');
      }
      throw new Error('Required columns not found in Daxko sheet');
    }
    
    // Get all data from Daxko sheet
    const daxkoData = daxkoSheet.getRange(2, 1, daxkoSheet.getLastRow() - 1, daxkoSheet.getLastColumn()).getValues();
    
    // Extract unique class combinations
    const classMap = new Map();
    
    daxkoData.forEach(row => {
      const program = row[programColIndex];
      const day = row[dayColIndex];
      const time = row[timeColIndex];
      const site = siteColIndex !== -1 ? row[siteColIndex] : '';
      
      // Skip if any required field is missing
      if (!program || !day || !time) return;
      
      const classKey = `${program}|${day}|${time}|${site}`;
      
      if (!classMap.has(classKey)) {
        classMap.set(classKey, {
          program: program,
          day: day,
          time: time,
          site: site,
          count: 1
        });
      } else {
        // Increment count for existing class
        const classInfo = classMap.get(classKey);
        classInfo.count++;
        classMap.set(classKey, classInfo);
      }
    });
    
    // Log class count
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Found ${classMap.size} unique classes`, 'INFO', 'setupClassSelector');
    }
    
    // Set up Classes sheet
    classesSheet.clear();
    
    // Set headers
    const headers = ['Select Class', 'Program/Stage', 'Day', 'Time', 'Location', 'Student Count', 'Instructor'];
    classesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    classesSheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add class options
    let rowIndex = 2;
    classMap.forEach(classInfo => {
      const rowData = [
        '',  // Selector cell will be converted to dropdown
        classInfo.program,
        classInfo.day,
        classInfo.time,
        classInfo.site,
        classInfo.count,
        ''   // Instructor column (to be filled by user or imported)
      ];
      
      classesSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      rowIndex++;
    });
    
    // Create data validation for the selector column
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Select'], true)
      .build();
    
    classesSheet.getRange(2, 1, rowIndex - 2, 1).setDataValidation(rule);
    
    // Set column widths for better visibility
    classesSheet.setColumnWidth(1, 100);  // Select Class
    classesSheet.setColumnWidth(2, 150);  // Program/Stage
    classesSheet.setColumnWidth(3, 100);  // Day
    classesSheet.setColumnWidth(4, 120);  // Time
    classesSheet.setColumnWidth(5, 120);  // Location
    classesSheet.setColumnWidth(6, 100);  // Student Count
    classesSheet.setColumnWidth(7, 150);  // Instructor
    
    // Freeze header row
    classesSheet.setFrozenRows(1);
    
    // Log success
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Class selector setup completed successfully', 'INFO', 'setupClassSelector');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'setupClassSelector', 
        'Failed to set up class selector. Please check your data and try again.');
    } else {
      Logger.log(`Class selector setup error: ${error.message}`);
      throw error; // Re-throw to caller
    }
    return false;
  }
}

/**
 * Updates the class selector in the Classes sheet.
 * This function can be called to refresh class data.
 * 
 * @return {boolean} Success status
 */
function updateClassSelector() {
  try {
    // Just call the setup function to refresh
    setupClassSelector();
    
    // Update instructor data if available
    try {
      const config = AdministrativeModule.getSystemConfiguration();
      if (config.sessionProgramsUrl) {
        importInstructorData(config.sessionProgramsUrl);
      }
    } catch (error) {
      // Log but continue
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Optional instructor data update error: ${error.message}`, 'WARNING', 'updateClassSelector');
      } else {
        Logger.log(`Optional instructor data update error: ${error.message}`);
      }
    }
    
    SpreadsheetApp.getUi().alert('Class Selector Updated', 'The class selector has been updated with the latest data.', SpreadsheetApp.getUi().ButtonSet.OK);
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'updateClassSelector', 
        'Failed to update class selector. Please try again or contact support.');
    } else {
      Logger.log(`Class selector update error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Update Failed', `Failed to update class selector: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Refreshes roster data from the source file.
 * 
 * @return {boolean} Success status
 */
function refreshRosterData() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Ask for confirmation
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      const confirmed = ErrorHandling.confirmAction(
        'Refresh Roster Data',
        'This will update the Daxko sheet with the latest data from the session roster file. Continue?'
      );
      
      if (!confirmed) return false;
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Refresh Roster Data',
        'This will update the Daxko sheet with the latest data from the session roster file. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      if (result !== ui.Button.YES) return false;
    }
    
    // Get configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    // Import roster data
    importSessionRosterData(config.sessionName, config.rosterFolderUrl);
    
    // Update class selector
    updateClassSelector();
    
    ui.alert('Roster Updated', 'The roster data has been refreshed successfully.', ui.ButtonSet.OK);
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'refreshRosterData', 
        'Failed to refresh roster data. Please check your configuration and try again.');
    } else {
      Logger.log(`Roster refresh error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Refresh Failed', `Failed to refresh roster data: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Imports instructor data from the Session Programs workbook.
 *
 * @param {string} sessionProgramsUrl - URL of the Session Programs workbook
 * @return {boolean} Success status
 */
function importInstructorData(sessionProgramsUrl) {
  // Log function entry
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage(`Importing instructor data from: ${sessionProgramsUrl}`, 'INFO', 'importInstructorData');
  }
  
  // If no URL is provided, log a warning and return
  if (!sessionProgramsUrl) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('No Session Programs URL provided. Skipping instructor data import.', 'WARNING', 'importInstructorData');
    }
    return true;
  }
  
  try {
    // Extract spreadsheet ID from URL
    let ssId;
    if (typeof GlobalFunctions.extractIdFromUrl === 'function') {
      ssId = GlobalFunctions.extractIdFromUrl(sessionProgramsUrl);
    } else {
      // Fallback to basic extraction
      ssId = extractIdFromUrl(sessionProgramsUrl);
    }
    
    if (!ssId) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Invalid Session Programs URL. Skipping instructor data import.', 'WARNING', 'importInstructorData');
      }
      return true; // Continue initialization even if URL is invalid
    }
    
    // Open the Session Programs workbook
    let programsWorkbook;
    if (typeof GlobalFunctions.safeGetSpreadsheetById === 'function') {
      programsWorkbook = GlobalFunctions.safeGetSpreadsheetById(ssId);
    } else {
      // Fallback to direct access
      programsWorkbook = SpreadsheetApp.openById(ssId);
    }
    
    if (!programsWorkbook) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Could not access Session Programs workbook. Skipping instructor data import.', 'WARNING', 'importInstructorData');
      }
      return true;
    }
    
    // Get the active sheet or the first sheet if multiple exist
    let programsSheet = programsWorkbook.getActiveSheet();
    if (!programsSheet) {
      const sheets = programsWorkbook.getSheets();
      if (sheets.length > 0) {
        programsSheet = sheets[0];
      } else {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage('No sheets found in the Session Programs workbook. Skipping instructor data import.', 'WARNING', 'importInstructorData');
        }
        return true; // Continue initialization even if no sheets are found
      }
    }
    
    // Get all data from Programs sheet
    const programsData = programsSheet.getDataRange().getValues();
    
    if (programsData.length <= 1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Session Programs sheet appears to be empty. Skipping instructor data import.', 'WARNING', 'importInstructorData');
      }
      return true; // Continue initialization even if sheet is empty
    }
    
    // Extract headers from the first row
    const headers = programsData[0];
    
    // Find relevant column indices with better matching function
    let programColIndex, dayColIndex, timeColIndex, locationColIndex, instructorColIndex;
    
    if (typeof GlobalFunctions.findColumnIndex === 'function') {
      programColIndex = GlobalFunctions.findColumnIndex(headers, ['program', 'class', 'stage', 'level']);
      dayColIndex = GlobalFunctions.findColumnIndex(headers, ['day', 'days', 'day(s)', 'weekday']);
      timeColIndex = GlobalFunctions.findColumnIndex(headers, ['time', 'session time', 'class time']);
      locationColIndex = GlobalFunctions.findColumnIndex(headers, ['location', 'site', 'facility']);
      instructorColIndex = GlobalFunctions.findColumnIndex(headers, ['instructor', 'teacher', 'staff', 'coach']);
    } else {
      // Fallback to simplified match
      programColIndex = findColumnIndexBasic(headers, ['program', 'class', 'stage', 'level']);
      dayColIndex = findColumnIndexBasic(headers, ['day', 'days', 'day(s)', 'weekday']);
      timeColIndex = findColumnIndexBasic(headers, ['time', 'session time', 'class time']);
      locationColIndex = findColumnIndexBasic(headers, ['location', 'site', 'facility']);
      instructorColIndex = findColumnIndexBasic(headers, ['instructor', 'teacher', 'staff', 'coach']);
    }
    
    // If essential columns aren't found, log the issue but continue
    if (programColIndex === -1 || dayColIndex === -1 || timeColIndex === -1 || instructorColIndex === -1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Required columns not found in Session Programs sheet. Headers found: ${headers.join(', ')}. Skipping instructor data import.`, 'WARNING', 'importInstructorData');
      }
      return true; // Continue initialization despite missing columns
    }
    
    // Create a map of class identifiers to instructor names
    const instructorMap = new Map();
    
    for (let i = 1; i < programsData.length; i++) {
      const row = programsData[i];
      
      // Skip empty rows
      if (!row[programColIndex] || !row[dayColIndex] || !row[timeColIndex]) {
        continue;
      }
      
      const program = row[programColIndex].toString().trim();
      const day = row[dayColIndex].toString().trim();
      const time = row[timeColIndex].toString().trim();
      const location = locationColIndex !== -1 ? row[locationColIndex].toString().trim() : '';
      const instructor = row[instructorColIndex] ? row[instructorColIndex].toString().trim() : '';
      
      if (instructor) {
        // Create a unique key for this class
        const classKey = `${program}|${day}|${time}|${location}`;
        instructorMap.set(classKey, instructor);
        
        // Also create alternative keys with slight variations
        // This helps match classes even if formatting is slightly different
        const classKeyNoSite = `${program}|${day}|${time}|`;
        instructorMap.set(classKeyNoSite, instructor);
      }
    }
    
    // Update the Classes sheet with instructor information
    const updated = updateClassInstructors(instructorMap);
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      if (updated) {
        ErrorHandling.logMessage('Instructor data imported and updated successfully', 'INFO', 'importInstructorData');
      } else {
        ErrorHandling.logMessage('No instructor data was updated', 'INFO', 'importInstructorData');
      }
    }
    
    return true;
  } catch (error) {
    // Log the error but don't throw it to allow initialization to continue
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Instructor data import warning: ${error.message}`, 'WARNING', 'importInstructorData');
    } else {
      Logger.log(`Instructor data import warning: ${error.message}`);
    }
    return true; // Return true to continue initialization
  }
}

/**
 * Updates the Classes sheet with instructor information.
 *
 * @param {Map} instructorMap - Map of class identifiers to instructor names
 * @return {boolean} Success status
 */
function updateClassInstructors(instructorMap) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Classes sheet not found. Skipping instructor update.', 'WARNING', 'updateClassInstructors');
      }
      return false;
    }
    
    // Get all data from Classes sheet
    const classesData = classesSheet.getDataRange().getValues();
    
    if (classesData.length <= 1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Classes sheet appears to be empty. Skipping instructor update.', 'WARNING', 'updateClassInstructors');
      }
      return false;
    }
    
    // Find column indices - IMPORTANT: These are 0-based indices
    const programColIndex = 1; // B column
    const dayColIndex = 2;     // C column
    const timeColIndex = 3;    // D column
    const locationColIndex = 4; // E column
    const instructorColIndex = 6; // G column
    
    // Update instructor information for each class
    let updatedCount = 0;
    
    for (let i = 1; i < classesData.length; i++) {
      const row = classesData[i];
      
      const program = row[programColIndex].toString().trim();
      const day = row[dayColIndex].toString().trim();
      const time = row[timeColIndex].toString().trim();
      const location = row[locationColIndex] ? row[locationColIndex].toString().trim() : '';
      
      // Create different key variations to increase match probability
      const classKey = `${program}|${day}|${time}|${location}`;
      const classKeyNoSite = `${program}|${day}|${time}|`;
      
      let instructor = null;
      
      // Try exact match first
      if (instructorMap.has(classKey)) {
        instructor = instructorMap.get(classKey);
      } 
      // Try without site if exact match fails
      else if (instructorMap.has(classKeyNoSite)) {
        instructor = instructorMap.get(classKeyNoSite);
      }
      // Try fuzzy matching as a last resort
      else {
        // Look through all keys for a close match
        for (const [key, value] of instructorMap.entries()) {
          const parts = key.split('|');
          if (parts.length < 3) continue;
          
          const keyProgram = parts[0];
          const keyDay = parts[1];
          const keyTime = parts[2];
          
          // Check for close matches with program, day and time
          if (keyProgram.includes(program) || program.includes(keyProgram)) {
            if (keyDay.includes(day) || day.includes(keyDay)) {
              if (keyTime.includes(time) || time.includes(keyTime)) {
                instructor = value;
                break;
              }
            }
          }
        }
      }
      
      // Update instructor in the Classes sheet if found
      if (instructor) {
        // FIXED: The column is already 0-indexed (G = 6), but row needs +1 for 1-indexing
        classesSheet.getRange(i + 1, instructorColIndex + 1).setValue(instructor);
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Updated instructor for "${program}" (${day}, ${time}) to "${instructor}"`, 'DEBUG', 'updateClassInstructors');
        }
        
        updatedCount++;
      }
    }
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Updated ${updatedCount} classes with instructor information`, 'INFO', 'updateClassInstructors');
    }
    
    return updatedCount > 0;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error updating class instructors: ${error.message}`, 'ERROR', 'updateClassInstructors');
    } else {
      Logger.log(`Error updating class instructors: ${error.message}`);
    }
    return false;
  }
}

/**
 * Pushes assessment data from class sheets to the YSL Swimmer Log.
 * 
 * @return {boolean} Success status
 */
function pushAssessmentsToSwimmerLog() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Ask for confirmation
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      const confirmed = ErrorHandling.confirmAction(
        'Push Assessments',
        'This will update the YSL Swimmer Log with the latest assessment data from class sheets. Continue?'
      );
      
      if (!confirmed) return false;
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Push Assessments',
        'This will update the YSL Swimmer Log with the latest assessment data from class sheets. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      if (result !== ui.Button.YES) return false;
    }
    
    // Get configuration
    const config = AdministrativeModule.getSystemConfiguration();
    const ssId = GlobalFunctions.extractIdFromUrl(config.swimmerRecordsUrl);
    
    if (!ssId) {
      throw new Error('Invalid Swimmer Log URL');
    }
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting push of assessments to Swimmer Log', 'INFO', 'pushAssessmentsToSwimmerLog');
    }
    
    // Open the Swimmer Log workbook
    let swimmerLog;
    if (typeof GlobalFunctions.safeGetSpreadsheetById === 'function') {
      swimmerLog = GlobalFunctions.safeGetSpreadsheetById(ssId);
    } else {
      // Fallback to direct access
      swimmerLog = SpreadsheetApp.openById(ssId);
    }
    
    if (!swimmerLog) {
      throw new Error('Could not open Swimmer Log workbook. Please check the URL and permissions.');
    }
    
    // Get the Swimmers sheet
    const swimmersSheet = swimmerLog.getSheetByName('Swimmers');
    
    if (!swimmersSheet) {
      throw new Error('Swimmers sheet not found in the Swimmer Log workbook');
    }
    
    // Get all swimmer data from YSL Swimmer Log
    const swimmerHeaders = swimmersSheet.getRange(1, 1, 1, swimmersSheet.getLastColumn()).getValues()[0];
    const swimmerData = swimmersSheet.getDataRange().getValues();
    
    // Find name and DOB column indices
    const nameColIndex = swimmerHeaders.indexOf('Name');
    const dobColIndex = swimmerHeaders.indexOf('DOB');
    
    if (nameColIndex === -1 || dobColIndex === -1) {
      throw new Error('Required columns not found in YSL Swimmer Log');
    }
    
    // Get all class sheets (sheets with Class_ prefix or _Assessment suffix)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let classSheets = sheets.filter(sheet => sheet.getName().startsWith('Class_'));
    
    // If no class sheets found, check if there are any assessment sheets
    if (classSheets.length === 0) {
      classSheets = sheets.filter(sheet => sheet.getName().includes('_Assessment'));
    }
    
    if (classSheets.length === 0) {
      throw new Error('No class assessment sheets found');
    }
    
    // Track results
    let updatedCount = 0;
    let errorCount = 0;
    let notFoundCount = 0;
    
    // Process each class sheet
    classSheets.forEach(classSheet => {
      try {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Processing sheet: ${classSheet.getName()}`, 'INFO', 'pushAssessmentsToSwimmerLog');
        }
        
        // Get assessment data from class sheet
        const classData = classSheet.getDataRange().getValues();
        const classHeaders = classData[0];
        
        // Find name and assessment columns
        const classNameColIndex = classHeaders.indexOf('Swimmer');
        const classDobColIndex = classHeaders.indexOf('DOB');
        
        if (classNameColIndex === -1) {
          throw new Error(`Required columns not found in sheet ${classSheet.getName()}`);
        }
        
        // Process each student in the class
        for (let i = 1; i < classData.length; i++) {
          const studentRow = classData[i];
          const studentName = studentRow[classNameColIndex];
          const studentDob = classDobColIndex !== -1 ? studentRow[classDobColIndex] : '';
          
          if (!studentName) continue;
          
          // Find the student in YSL Swimmer Log
          let studentFound = false;
          let studentRowIndex = -1;
          
          for (let j = 1; j < swimmerData.length; j++) {
            const swimmerName = swimmerData[j][nameColIndex];
            const swimmerDob = swimmerData[j][dobColIndex];
            
            // Match by name and DOB if available, otherwise just by name
            if (swimmerName === studentName && 
                (classDobColIndex === -1 || !studentDob || !swimmerDob || studentDob === swimmerDob)) {
              studentFound = true;
              studentRowIndex = j;
              break;
            }
          }
          
          if (studentFound) {
            // Update assessment data for the found student
            let updatedCells = 0;
            
            for (let c = 0; c < classHeaders.length; c++) {
              const header = classHeaders[c];
              
              // Skip non-assessment columns
              if (['Swimmer', 'DOB', 'Age', 'Gender', 'Class', 'Instructor', 'Notes'].includes(header)) {
                continue;
              }
              
              const swimmerColIndex = swimmerHeaders.indexOf(header);
              
              if (swimmerColIndex !== -1 && studentRow[c] !== '') {
                // Update the cell in YSL Swimmer Log
                swimmersSheet.getRange(studentRowIndex + 1, swimmerColIndex + 1).setValue(studentRow[c]);
                updatedCells++;
              }
            }
            
            if (updatedCells > 0) {
              updatedCount++;
              
              if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
                ErrorHandling.logMessage(`Updated ${updatedCells} cells for student: ${studentName}`, 'DEBUG', 'pushAssessmentsToSwimmerLog');
              }
            }
          } else {
            notFoundCount++;
            
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`Student not found: ${studentName}`, 'WARNING', 'pushAssessmentsToSwimmerLog');
            }
          }
        }
      } catch (error) {
        errorCount++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error processing sheet ${classSheet.getName()}: ${error.message}`, 'ERROR', 'pushAssessmentsToSwimmerLog');
        } else {
          Logger.log(`Error processing sheet ${classSheet.getName()}: ${error.message}`);
        }
      }
    });
    
    // Log results
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Assessment push completed: ${updatedCount} students updated, ${notFoundCount} not found, ${errorCount} sheets with errors`, 'INFO', 'pushAssessmentsToSwimmerLog');
    }
    
    // Show results
    ui.alert(
      'Push Complete',
      `Assessment data pushed to YSL Swimmer Log:\n` +
      `- ${updatedCount} students updated\n` +
      `- ${notFoundCount} students not found\n` +
      `- ${errorCount} sheets with errors\n\n` +
      `See logs for details.`,
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'pushAssessmentsToSwimmerLog', 
        'Failed to push assessment data. Please check your configuration and try again.');
    } else {
      Logger.log(`Assessment push error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Push Failed', `Failed to push assessment data: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Pulls the latest assessment criteria from YSL Swimmer Log.
 * 
 * @return {boolean} Success status
 */
function pullAssessmentCriteria() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Ask for confirmation
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      const confirmed = ErrorHandling.confirmAction(
        'Update Assessment Criteria',
        'This will update the assessment criteria from the YSL Swimmer Log. Continue?'
      );
      
      if (!confirmed) return false;
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Update Assessment Criteria',
        'This will update the assessment criteria from the YSL Swimmer Log. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      if (result !== ui.Button.YES) return false;
    }
    
    // Get configuration and import assessment criteria
    const config = AdministrativeModule.getSystemConfiguration();
    importAssessmentCriteria(config.swimmerRecordsUrl);
    
    ui.alert('Criteria Updated', 'Assessment criteria have been updated successfully.', ui.ButtonSet.OK);
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'pullAssessmentCriteria', 
        'Failed to update assessment criteria. Please check your configuration and try again.');
    } else {
      Logger.log(`Assessment criteria update error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Update Failed', `Failed to update assessment criteria: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Gets roster data for a specific class.
 * 
 * @param {string} program - The program/stage
 * @param {string} day - The day of the week
 * @param {string} time - The class time
 * @param {string} site - The class location
 * @return {Object} Object containing roster data and headers
 */
function getRosterForClass(program, day, time, site) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName('Daxko');
    
    if (!daxkoSheet) {
      throw new Error('Daxko sheet not found');
    }
    
    // Get all data from Daxko sheet
    const daxkoData = daxkoSheet.getDataRange().getValues();
    const headers = daxkoData[0];
    
    // Log the search parameters
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Searching for roster with Program="${program}", Day="${day}", Time="${time}", Site="${site}"`, 'DEBUG', 'getRosterForClass');
    }
    
    // Find the column indices for relevant fields
    let programColIndex, dayColIndex, timeColIndex, siteColIndex, firstNameColIndex, lastNameColIndex, dobColIndex, genderColIndex;
    
    // Use the findColumnIndex function if available, otherwise use indexOf
    if (typeof GlobalFunctions.findColumnIndex === 'function') {
      programColIndex = GlobalFunctions.findColumnIndex(headers, ['program', 'class', 'stage', 'level']);
      dayColIndex = GlobalFunctions.findColumnIndex(headers, ['day', 'days', 'day(s)', 'weekday']);
      timeColIndex = GlobalFunctions.findColumnIndex(headers, ['time', 'session time', 'class time']);
      siteColIndex = GlobalFunctions.findColumnIndex(headers, ['location', 'site', 'facility']);
      firstNameColIndex = GlobalFunctions.findColumnIndex(headers, ['first name', 'firstname', 'given name']);
      lastNameColIndex = GlobalFunctions.findColumnIndex(headers, ['last name', 'lastname', 'surname', 'family name']);
      dobColIndex = GlobalFunctions.findColumnIndex(headers, ['dob', 'birth date', 'date of birth', 'birthdate']);
      genderColIndex = GlobalFunctions.findColumnIndex(headers, ['gender', 'sex']);
    } else {
      // Fallback to basic indexOf
      programColIndex = headers.indexOf('Program');
      dayColIndex = headers.indexOf('Day(s) of Week');
      timeColIndex = headers.indexOf('Session Time');
      siteColIndex = headers.indexOf('Site');
      firstNameColIndex = headers.indexOf('First Name');
      lastNameColIndex = headers.indexOf('Last Name');
      dobColIndex = headers.indexOf('DOB');
      genderColIndex = headers.indexOf('Gender');
    }
    
    if (programColIndex === -1 || dayColIndex === -1 || timeColIndex === -1 ||
        firstNameColIndex === -1 || lastNameColIndex === -1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Required columns not found in Daxko sheet. Headers: ${headers.join(', ')}`, 'ERROR', 'getRosterForClass');
      }
      throw new Error('Required columns not found in Daxko sheet');
    }
    
    // Filter data for the specified class
    const roster = [];
    
    for (let i = 1; i < daxkoData.length; i++) {
      const row = daxkoData[i];
      
      // Normalize data for matching
      const rowProgram = String(row[programColIndex]).trim();
      const rowDay = String(row[dayColIndex]).trim();
      const rowTime = String(row[timeColIndex]).trim();
      const rowSite = siteColIndex !== -1 ? String(row[siteColIndex]).trim() : '';
      
      // More flexible matching
      const programMatch = rowProgram === program || rowProgram.includes(program) || program.includes(rowProgram);
      const dayMatch = rowDay === day || rowDay.replace('.', '') === day.replace('.', '');
      const timeMatch = rowTime === time || rowTime.includes(time) || time.includes(rowTime);
      const siteMatch = siteColIndex === -1 || site === '' || rowSite === '' || rowSite === site;
      
      if (programMatch && dayMatch && timeMatch && siteMatch) {
        // Construct student record
        const student = {
          name: `${row[firstNameColIndex]} ${row[lastNameColIndex]}`,
          firstName: row[firstNameColIndex],
          lastName: row[lastNameColIndex],
          dob: dobColIndex !== -1 ? row[dobColIndex] : '',
          gender: genderColIndex !== -1 ? row[genderColIndex] : '',
          rowData: row,
          rowIndex: i
        };
        
        roster.push(student);
      }
    }
    
    // Log the results
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Found ${roster.length} students for class: ${program} (${day}, ${time})`, 'INFO', 'getRosterForClass');
    }
    
    return {
      roster: roster,
      headers: headers
    };
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting roster for class: ${error.message}`, 'ERROR', 'getRosterForClass');
    } else {
      Logger.log(`Error getting roster for class: ${error.message}`);
    }
    throw error; // Re-throw to caller
  }
}

/**
 * Gets assessment criteria for a specific stage.
 * 
 * @param {string} stage - The swim stage/level
 * @return {Object} Object containing criteria data
 */
function getAssessmentCriteriaForStage(stage) {
  try {
    // Log the requested stage
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Getting assessment criteria for stage: ${stage}`, 'DEBUG', 'getAssessmentCriteriaForStage');
    }
    
    // Convert stage name to expected format (e.g., "Stage 1" -> "S1")
    let stageCode = '';
    
    if (stage.includes('Stage')) {
      // Extract number or letter from stage name
      const match = stage.match(/Stage\s+([A-Za-z0-9]+)/);
      if (match && match[1]) {
        stageCode = 'S' + match[1];
      }
    } else if (stage === 'Private Swim Lessons') {
      // For private lessons, determine stage from instructor input or default to comprehensive criteria
      stageCode = 'private';
    } else {
      // Direct mapping if stage is already in code format
      stageCode = stage;
    }
    
    // Get stored criteria
    const scriptProperties = PropertiesService.getScriptProperties();
    const criteriaJson = scriptProperties.getProperty('assessmentCriteria');
    const headersJson = scriptProperties.getProperty('criteriaHeaders');
    
    if (!criteriaJson || !headersJson) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Assessment criteria not found in script properties', 'ERROR', 'getAssessmentCriteriaForStage');
      }
      throw new Error('Assessment criteria not found. Please run initialization again.');
    }
    
    const criteria = JSON.parse(criteriaJson);
    const headers = JSON.parse(headersJson);
    
    // Find criteria for the specified stage
    const stageCriteria = [];
    
    for (let i = 0; i < criteria.length; i++) {
      const header = headers[i];
      const value = criteria[i];
      
      // Check if header starts with the stage code
      if (header && (header.startsWith(stageCode) || stageCode === 'private')) {
        stageCriteria.push({
          header: header,
          value: value
        });
      }
    }
    
    // Log the results
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Found ${stageCriteria.length} criteria items for stage: ${stage} (code: ${stageCode})`, 'INFO', 'getAssessmentCriteriaForStage');
    }
    
    return {
      stageCode: stageCode,
      criteria: stageCriteria,
      allCriteria: criteria,
      allHeaders: headers
    };
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting assessment criteria: ${error.message}`, 'ERROR', 'getAssessmentCriteriaForStage');
    } else {
      Logger.log(`Error getting assessment criteria: ${error.message}`);
    }
    throw error; // Re-throw to caller
  }
}

/**
 * Fallback function to find column index by possible names
 * Used if the GlobalFunctions version is not available
 * 
 * @param {Array} headers - Array of header names
 * @param {Array} possibleNames - Array of possible names for the column
 * @return {number} The index of the found column, or -1 if not found
 */
function findColumnIndexBasic(headers, possibleNames) {
  // First try exact matches (case-insensitive)
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] ? headers[i].toString().toLowerCase() : '';
    
    for (const name of possibleNames) {
      if (header === name.toLowerCase()) {
        return i; // Exact match
      }
    }
  }
  
  // Then try partial matches (case-insensitive)
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] ? headers[i].toString().toLowerCase() : '';
    
    for (const name of possibleNames) {
      if (header.includes(name.toLowerCase()) || name.toLowerCase().includes(header)) {
        return i; // Partial match
      }
    }
  }
  
  return -1; // No match found
}

/**
 * Utility function to extract ID from a Google Drive URL.
 * Fallback if GlobalFunctions version is not available.
 * 
 * @param {string} url - The Google Drive URL
 * @return {string|null} The extracted ID, or null if not found
 */
function extractIdFromUrl(url) {
  if (!url) return null;
  
  // Extract folder ID from various URL formats
  const patterns = [
    /\/folders\/([a-zA-Z0-9-_]+)/,         // Drive folder URL
    /\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/, // Spreadsheet URL
    /id=([a-zA-Z0-9-_]+)/,                 // URL parameter format
    /^([a-zA-Z0-9-_]+)$/                   // Direct ID
  ];
  
  for (const pattern of patterns) {
    const match = url.match(pattern);
    if (match && match[1]) {
      return match[1];
    }
  }
  
  return null;
}

/**
 * Reports the assessment criteria values found when pulling from Swimmer Log.
 * This function can be called after pullAssessmentCriteria to verify data.
 * 
 * @return {boolean} Success status
 */
/**
 * Reports the assessment criteria values found when pulling from Swimmer Log.
 * This function can be called after pullAssessmentCriteria to verify data.
 * 
 * @return {boolean} Success status
 */
function reportAssessmentCriteria() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get stored criteria from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const criteriaJson = scriptProperties.getProperty('assessmentCriteria');
    const headersJson = scriptProperties.getProperty('criteriaHeaders');
    
    if (!criteriaJson || !headersJson) {
      ui.alert(
        'No Assessment Criteria Found',
        'No assessment criteria found in script properties. Please run "Pull Assessment Criteria" first.',
        ui.ButtonSet.OK
      );
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('No assessment criteria found in script properties', 'ERROR', 'reportAssessmentCriteria');
      }
      
      return false;
    }
    
    // Parse the criteria and headers
    const criteria = JSON.parse(criteriaJson);
    const headers = JSON.parse(headersJson);
    
    // Create a report sheet or clear existing one
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let reportSheet = ss.getSheetByName('CriteriaReport');
    
    if (reportSheet) {
      reportSheet.clear();
    } else {
      reportSheet = ss.insertSheet('CriteriaReport');
    }
    
    // Set up headers
    reportSheet.getRange(1, 1, 1, 3).setValues([['Header', 'Criteria Value', 'Stage Code']]);
    reportSheet.getRange(1, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Collect all data rows first
    const rowsData = [];
    
    // Group headers and values by stage code
    const stageGroups = {};
    
    // First, identify all stage codes
    for (let i = 0; i < headers.length; i++) {
      if (!headers[i]) continue;
      
      // Extract stage code if header has correct format (e.g., S1, S2, S3, etc.)
      let stageCode = '';
      if (headers[i].match(/^[A-Za-z][0-9A-Za-z]/)) {
        stageCode = headers[i].substring(0, 2);
        
        if (!stageGroups[stageCode]) {
          stageGroups[stageCode] = [];
        }
        
        stageGroups[stageCode].push({
          index: i,
          header: headers[i],
          value: criteria[i] || ''
        });
      }
    }
    
    // Now process the data by stage
    for (const stageCode in stageGroups) {
      const stageData = stageGroups[stageCode];
      
      // Add a stage header row
      rowsData.push([`${stageCode} Criteria`, '', '']);
      
      // Add all criteria for this stage
      for (const item of stageData) {
        rowsData.push([
          item.header,
          item.value,
          stageCode
        ]);
      }
      
      // Add separator row
      rowsData.push(['', '', '']);
    }
    
    // Add non-stage-specific headers at the end
    const otherHeaders = [];
    for (let i = 0; i < headers.length; i++) {
      if (!headers[i]) continue;
      
      // If not a stage-specific header, add to other headers
      if (!headers[i].match(/^[A-Za-z][0-9A-Za-z]/)) {
        otherHeaders.push({
          header: headers[i],
          value: criteria[i] || ''
        });
      }
    }
    
    if (otherHeaders.length > 0) {
      rowsData.push(['Other Criteria', '', '']);
      for (const item of otherHeaders) {
        rowsData.push([
          item.header,
          item.value,
          ''
        ]);
      }
    }
    
    // Write all data to the sheet
    if (rowsData.length > 0) {
      reportSheet.getRange(2, 1, rowsData.length, 3).setValues(rowsData);
      
      // Apply formatting to stage header rows
      for (let i = 0; i < rowsData.length; i++) {
        const row = i + 2; // Start from row 2 (after column headers)
        
        if (rowsData[i][0].endsWith('Criteria')) {
          // This is a stage header row
          reportSheet.getRange(row, 1, 1, 3)
            .setBackground('#4285F4')
            .setFontColor('white')
            .setFontWeight('bold');
        } else if (rowsData[i][0] === '' && rowsData[i][1] === '' && rowsData[i][2] === '') {
          // This is a separator row
          reportSheet.getRange(row, 1, 1, 3).setBackground('#E0E0E0');
        } else {
          // Regular row, apply alternating colors
          if (i % 2 === 1) {
            reportSheet.getRange(row, 1, 1, 3).setBackground('#f5f5f5');
          }
        }
      }
    }
    
    // Format sheet
    reportSheet.autoResizeColumn(1);
    reportSheet.autoResizeColumn(2);
    reportSheet.autoResizeColumn(3);
    reportSheet.setFrozenRows(1);
    
    // Activate the report sheet
    reportSheet.activate();
    
    // Count stats
    let totalCriteria = 0;
    for (const stageCode in stageGroups) {
      totalCriteria += stageGroups[stageCode].length;
    }
    totalCriteria += otherHeaders.length;
    
    // Show summary
    ui.alert(
      'Assessment Criteria Report',
      `Report generated successfully.\n\n` +
      `Total criteria found: ${totalCriteria}\n` +
      `Stage codes found: ${Object.keys(stageGroups).join(', ')}\n\n` +
      `The report shows all criteria for each stage, including those between pass values.\n` +
      `See the CriteriaReport sheet for details.`,
      ui.ButtonSet.OK
    );
    
    // Log the successful report
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Assessment criteria report generated with ${totalCriteria} criteria across ${Object.keys(stageGroups).length} stages`, 'INFO', 'reportAssessmentCriteria');
    }
    
    return true;
  } catch (error) {
    // Handle any unexpected errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'reportAssessmentCriteria', 
        'Failed to generate assessment criteria report. Please try again or contact support.');
    } else {
      Logger.log(`Error generating assessment criteria report: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to generate report: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Diagnostic function to report raw assessment criteria data.
 * This will show exactly what was stored in the script properties.
 * 
 * @return {boolean} Success status
 */
function diagnoseCriteriaImport() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get stored criteria directly from script properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const criteriaJson = scriptProperties.getProperty('assessmentCriteria');
    const headersJson = scriptProperties.getProperty('criteriaHeaders');
    
    if (!criteriaJson || !headersJson) {
      ui.alert(
        'No Assessment Criteria Found',
        'No assessment criteria found in script properties. Please run "Pull Assessment Criteria" first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Parse the JSON strings to get the actual data
    const criteria = JSON.parse(criteriaJson);
    const headers = JSON.parse(headersJson);
    
    // Create a diagnostic sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let diagSheet = ss.getSheetByName('CriteriaDiagnostic');
    
    if (diagSheet) {
      diagSheet.clear();
    } else {
      diagSheet = ss.insertSheet('CriteriaDiagnostic');
    }
    
    // Add metadata
    diagSheet.getRange(1, 1).setValue('Assessment Criteria Diagnostic');
    diagSheet.getRange(2, 1).setValue(`Headers Length: ${headers.length}`);
    diagSheet.getRange(3, 1).setValue(`Criteria Length: ${criteria.length}`);
    diagSheet.getRange(4, 1).setValue(`Generated: ${new Date()}`);
    
    // Set up column headers
    diagSheet.getRange(6, 1, 1, 3).setValues([['Index', 'Header Value', 'Criteria Value']]);
    diagSheet.getRange(6, 1, 1, 3)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Prepare the raw data
    const rowsData = [];
    for (let i = 0; i < Math.max(headers.length, criteria.length); i++) {
      rowsData.push([
        i,
        i < headers.length ? headers[i] : 'MISSING',
        i < criteria.length ? criteria[i] : 'MISSING'
      ]);
    }
    
    // Write all data to the sheet
    if (rowsData.length > 0) {
      diagSheet.getRange(7, 1, rowsData.length, 3).setValues(rowsData);
    }
    
    // Format sheet
    diagSheet.autoResizeColumn(1);
    diagSheet.autoResizeColumn(2);
    diagSheet.autoResizeColumn(3);
    diagSheet.setFrozenRows(6);
    
    // Activate the diagnostic sheet
    diagSheet.activate();
    
    // Count stage-specific criteria
    let stageSpecificCount = 0;
    for (let i = 0; i < headers.length; i++) {
      if (headers[i] && headers[i].match(/^[A-Za-z][0-9A-Za-z]/)) {
        stageSpecificCount++;
      }
    }
    
    // Show summary
    ui.alert(
      'Criteria Diagnostic Complete',
      `Diagnostic report generated.\n\n` +
      `Total headers: ${headers.length}\n` +
      `Total criteria values: ${criteria.length}\n` +
      `Stage-specific headers: ${stageSpecificCount}\n\n` +
      `This report shows the RAW data stored in script properties.\n` +
      `See the CriteriaDiagnostic sheet for detailed analysis.`,
      ui.ButtonSet.OK
    );
    
    // Add the diagnostic function to the logging
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Assessment criteria diagnostic completed', 'INFO', 'diagnoseCriteriaImport');
    }
    
    return true;
  } catch (error) {
    // Error handling
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'diagnoseCriteriaImport', 
        'Failed to generate criteria diagnostic. Please try again or contact support.');
    } else {
      Logger.log(`Error generating criteria diagnostic: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to generate diagnostic: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

// Make functions available to other modules
const DataIntegrationModule = {
  initializeDataStructures: initializeDataStructures,
  updateClassSelector: updateClassSelector,
  refreshRosterData: refreshRosterData,
  pushAssessmentsToSwimmerLog: pushAssessmentsToSwimmerLog,
  pullAssessmentCriteria: pullAssessmentCriteria,
  getRosterForClass: getRosterForClass,
  getAssessmentCriteriaForStage: getAssessmentCriteriaForStage,
  importInstructorData: importInstructorData,
  importSessionRosterData: importSessionRosterData,
  reportAssessmentCriteria: reportAssessmentCriteria,
  diagnoseCriteriaImport: diagnoseCriteriaImport
};