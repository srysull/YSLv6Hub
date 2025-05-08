/**
 * YSL Hub v2 Blank Sheet Initializer
 * 
 * This module provides functions to create a complete system structure from a blank spreadsheet.
 * It handles the creation of all required sheets, basic formatting, and guides the user through
 * the initial configuration process.
 * 
 * @author Claude Code
 * @version 1.0
 * @date 2025-05-05
 */

/**
 * Initializes a completely blank spreadsheet with all required sheets and basic structure
 * This function should be the first step when starting with a new blank spreadsheet
 */
function initializeBlankSpreadsheet() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Log the start of the process
    Logger.log('Starting blank spreadsheet initialization');
    
    // Confirm with the user
    const result = ui.alert(
      'Initialize YSL Hub',
      'This will create all required sheets and structure for a new YSL Hub system. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return false;
    }
    
    // Create the base structure
    createBaseStructure(ss);
    
    // Show success message
    ui.alert(
      'Basic Structure Created',
      'The basic YSL Hub structure has been created. You will now be guided through the configuration process.',
      ui.ButtonSet.OK
    );
    
    // Launch the initialization dialog
    if (AdministrativeModule && typeof AdministrativeModule.showInitializationDialog === 'function') {
      AdministrativeModule.showInitializationDialog();
    } else {
      // Fallback if function not available
      ui.alert(
        'Next Steps',
        'Please select "Initialize System" from the YSL Hub menu to complete the setup.',
        ui.ButtonSet.OK
      );
    }
    
    return true;
  } catch (error) {
    // Basic error handling since error handling system might not be initialized yet
    Logger.log(`Error initializing blank spreadsheet: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'Initialization Error',
      `An error occurred during initialization: ${error.message}. Please try again or contact support.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}

/**
 * Creates the base sheet structure for a new YSL Hub system
 * @param {Spreadsheet} ss - The active spreadsheet
 */
function createBaseStructure(ss) {
  // Get existing sheets
  const existingSheets = ss.getSheets();
  
  // Required sheet names and their creation order
  const requiredSheets = [
    'Assumptions',
    'Classes',
    'Roster',
    'Assessments',
    'Announcements',
    'Reports',
    'SystemLog'
  ];
  
  // Configuration for each sheet
  const sheetConfig = {
    'Assumptions': {
      headers: ['Configuration Parameter', 'Value'],
      columnWidths: [250, 400]
    },
    'Classes': {
      headers: ['Select Class', 'Program', 'Day', 'Time', 'Location', 'Count', 'Instructor', 'Notes'],
      columnWidths: [100, 200, 100, 100, 150, 70, 150, 300]
    },
    'Roster': {
      headers: ['Class ID', 'Student Name', 'Age', 'Parent/Guardian', 'Email', 'Phone', 'Special Notes'],
      columnWidths: [100, 200, 70, 200, 200, 150, 300]
    },
    'Assessments': {
      headers: ['Student ID', 'Student Name', 'Class', 'Skill', 'Rating', 'Comments', 'Date'],
      columnWidths: [100, 200, 150, 200, 100, 300, 120]
    },
    'Announcements': {
      headers: ['Class ID', 'Program', 'Day', 'Time', 'Instructor', 'Recipients', 'Subject', 'Message', 'Status', 'Sent Date'],
      columnWidths: [100, 150, 100, 100, 150, 150, 250, 400, 100, 120]
    },
    'Reports': {
      headers: ['Report Type', 'Date Generated', 'Class', 'Status', 'Count', 'Notes'],
      columnWidths: [150, 120, 200, 100, 70, 300]
    }
  };
  
  // Map for renaming existing sheets vs creating new ones
  const sheetMap = {};
  
  // If there are existing sheets, map the first one to Assumptions and create others
  if (existingSheets.length > 0) {
    sheetMap[requiredSheets[0]] = existingSheets[0];
    existingSheets[0].setName(requiredSheets[0]);
  }
  
  // Create all required sheets that don't already exist
  requiredSheets.forEach((sheetName) => {
    if (!sheetMap[sheetName]) {
      const newSheet = ss.insertSheet(sheetName);
      sheetMap[sheetName] = newSheet;
    }
  });
  
  // Format each sheet
  Object.keys(sheetConfig).forEach((sheetName) => {
    const sheet = sheetMap[sheetName];
    const config = sheetConfig[sheetName];
    
    // Set headers
    if (config.headers && config.headers.length > 0) {
      sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
      
      // Format header row
      sheet.getRange(1, 1, 1, config.headers.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3')
        .setHorizontalAlignment('center');
      
      // Freeze header row
      sheet.setFrozenRows(1);
    }
    
    // Set column widths
    if (config.columnWidths && config.columnWidths.length > 0) {
      config.columnWidths.forEach((width, index) => {
        sheet.setColumnWidth(index + 1, width);
      });
    }
  });
  
  // Hide SystemLog sheet
  if (sheetMap['SystemLog']) {
    sheetMap['SystemLog'].hideSheet();
  }
  
  // Additional formatting for Classes sheet
  if (sheetMap['Classes']) {
    // Add dropdown for Select Class column
    const classesSheet = sheetMap['Classes'];
    const selectRange = classesSheet.getRange("A2:A100");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Select', 'Exclude'], true)
      .build();
    selectRange.setDataValidation(rule);
  }
  
  // Additional formatting for Announcements sheet
  if (sheetMap['Announcements']) {
    // Add dropdown for Status column
    const announcementsSheet = sheetMap['Announcements'];
    const statusRange = announcementsSheet.getRange("I2:I100");
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Draft', 'Ready', 'Sent', 'Failed'], true)
      .build();
    statusRange.setDataValidation(rule);
  }
  
  // Set Assumptions sheet as active
  if (sheetMap['Assumptions']) {
    sheetMap['Assumptions'].activate();
  }
  
  Logger.log('Basic sheet structure created successfully');
}

/**
 * Adds a menu item for blank spreadsheet initialization
 */
function addInitializerMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('YSL Hub Setup')
      .addItem('Initialize Blank Spreadsheet', 'initializeBlankSpreadsheet')
      .addToUi();
    
    return true;
  } catch (error) {
    Logger.log(`Error adding initializer menu: ${error.message}`);
    return false;
  }
}