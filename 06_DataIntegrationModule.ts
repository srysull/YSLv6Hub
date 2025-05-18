/**
 * YSL Hub v2 Data Integration Module
 * 
 * This module handles the integration of data between various components of the
 * YSL Hub system, including roster data, assessments, and instructor information.
 * It provides functions for importing, exporting, and synchronizing data.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

// Sheet names for data storage
const DATA_SHEETS = {
  CLASSES: 'Classes',
  ROSTER: 'Roster',
  CRITERIA: 'AssessmentCriteria',
  INSTRUCTORS: 'Instructors'
};

// Column indices for class data
const CLASS_COLUMNS = {
  CLASS_ID: 0,       // A
  CLASS_NAME: 1,     // B
  CLASS_LEVEL: 2,    // C
  INSTRUCTOR: 3,     // D
  DAY: 4,            // E
  TIME: 5,           // F
  START_DATE: 6,     // G
  END_DATE: 7,       // H
  LOCATION: 8,       // I
  NUM_STUDENTS: 9,   // J
  CAPACITY: 10,      // K
  STATUS: 11,        // L
  NOTES: 12          // M
};

// Column indices for roster data
const ROSTER_COLUMNS = {
  STUDENT_ID: 0,     // A
  CLASS_ID: 1,       // B
  FIRST_NAME: 2,     // C
  LAST_NAME: 3,      // D
  AGE: 4,            // E
  LEVEL: 5,          // F
  NOTES: 6,          // G
  PARENT_EMAIL: 7,   // H
  PHONE: 8,          // I
  STATUS: 9,         // J
  ASSESSMENT_DATE: 10 // K
};

/**
 * Initialize data structures for the system
 * Creates necessary sheets and populates with initial data
 * 
 * @param config - The system configuration values
 */
function initializeDataStructures(config) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Initializing data structures', 'INFO', 'initializeDataStructures');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create Classes sheet if it doesn't exist
    createClassesSheet(ss);
    
    // Create Roster sheet if it doesn't exist
    createRosterSheet(ss);
    
    // Create AssessmentCriteria sheet if it doesn't exist
    createAssessmentCriteriaSheet(ss);
    
    // Create Instructors sheet if it doesn't exist
    createInstructorsSheet(ss);
    
    // Import initial data from session programs workbook
    if (config.sessionProgramsUrl) {
      importInstructorData(config.sessionProgramsUrl);
    }
    
    // Import assessment criteria from swimmer records workbook
    if (config.swimmerRecordsUrl) {
      pullAssessmentCriteria(config.swimmerRecordsUrl);
    }
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Data structures initialized successfully', 'INFO', 'initializeDataStructures');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'initializeDataStructures', 
        'Error initializing data structures. Some features may not work properly.');
    } else {
      Logger.log(`Error initializing data structures: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Initialization Error',
        `Failed to initialize data structures: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Creates the Classes sheet with appropriate headers and formatting
 * 
 * @param ss - The active spreadsheet
 * @returns The created or existing Classes sheet
 */
function createClassesSheet(ss) {
  let sheet = ss.getSheetByName(DATA_SHEETS.CLASSES);
  
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEETS.CLASSES);
    
    // Set up headers
    const headers = [
      'Class ID', 'Class Name', 'Level', 'Instructor', 'Day', 'Time', 
      'Start Date', 'End Date', 'Location', '# Students', 'Capacity', 'Status', 'Notes'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Set column widths for better visibility
    sheet.setColumnWidth(1, 80);   // Class ID
    sheet.setColumnWidth(2, 200);  // Class Name
    sheet.setColumnWidth(3, 100);  // Level
    sheet.setColumnWidth(4, 150);  // Instructor
    sheet.setColumnWidth(5, 100);  // Day
    sheet.setColumnWidth(6, 100);  // Time
    sheet.setColumnWidth(7, 120);  // Start Date
    sheet.setColumnWidth(8, 120);  // End Date
    sheet.setColumnWidth(9, 100);  // Location
    sheet.setColumnWidth(10, 100); // # Students
    sheet.setColumnWidth(11, 100); // Capacity
    sheet.setColumnWidth(12, 100); // Status
    sheet.setColumnWidth(13, 300); // Notes
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Create data validation for the Status column
    const statusRange = sheet.getRange(2, CLASS_COLUMNS.STATUS + 1, 100, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Cancelled', 'Completed', 'Pending'], true)
      .build();
    statusRange.setDataValidation(statusRule);
  }
  
  return sheet;
}

/**
 * Creates the Roster sheet with appropriate headers and formatting
 * 
 * @param ss - The active spreadsheet
 * @returns The created or existing Roster sheet
 */
function createRosterSheet(ss) {
  let sheet = ss.getSheetByName(DATA_SHEETS.ROSTER);
  
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEETS.ROSTER);
    
    // Set up headers
    const headers = [
      'Student ID', 'Class ID', 'First Name', 'Last Name', 'Age', 'Level', 
      'Notes', 'Parent Email', 'Phone', 'Status', 'Assessment Date'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Set column widths for better visibility
    sheet.setColumnWidth(1, 100);  // Student ID
    sheet.setColumnWidth(2, 100);  // Class ID
    sheet.setColumnWidth(3, 120);  // First Name
    sheet.setColumnWidth(4, 120);  // Last Name
    sheet.setColumnWidth(5, 60);   // Age
    sheet.setColumnWidth(6, 100);  // Level
    sheet.setColumnWidth(7, 250);  // Notes
    sheet.setColumnWidth(8, 200);  // Parent Email
    sheet.setColumnWidth(9, 120);  // Phone
    sheet.setColumnWidth(10, 100); // Status
    sheet.setColumnWidth(11, 120); // Assessment Date
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Create data validation for the Status column
    const statusRange = sheet.getRange(2, ROSTER_COLUMNS.STATUS + 1, 100, 1);
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Active', 'Waitlisted', 'Withdrawn', 'Completed'], true)
      .build();
    statusRange.setDataValidation(statusRule);
  }
  
  return sheet;
}

/**
 * Creates the AssessmentCriteria sheet with appropriate headers and formatting
 * 
 * @param ss - The active spreadsheet
 * @returns The created or existing AssessmentCriteria sheet
 */
function createAssessmentCriteriaSheet(ss) {
  let sheet = ss.getSheetByName(DATA_SHEETS.CRITERIA);
  
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEETS.CRITERIA);
    
    // Set up headers
    const headers = [
      'Level', 'Category', 'Skill', 'Description', 'Proficiency Required'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Set column widths for better visibility
    sheet.setColumnWidth(1, 100);  // Level
    sheet.setColumnWidth(2, 150);  // Category
    sheet.setColumnWidth(3, 200);  // Skill
    sheet.setColumnWidth(4, 300);  // Description
    sheet.setColumnWidth(5, 150);  // Proficiency Required
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Creates the Instructors sheet with appropriate headers and formatting
 * 
 * @param ss - The active spreadsheet
 * @returns The created or existing Instructors sheet
 */
function createInstructorsSheet(ss) {
  let sheet = ss.getSheetByName(DATA_SHEETS.INSTRUCTORS);
  
  if (!sheet) {
    sheet = ss.insertSheet(DATA_SHEETS.INSTRUCTORS);
    
    // Set up headers
    const headers = [
      'Instructor ID', 'First Name', 'Last Name', 'Email', 'Phone', 
      'Certifications', 'Availability', 'Notes'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    
    // Format header row
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Set column widths for better visibility
    sheet.setColumnWidth(1, 100);  // Instructor ID
    sheet.setColumnWidth(2, 120);  // First Name
    sheet.setColumnWidth(3, 120);  // Last Name
    sheet.setColumnWidth(4, 200);  // Email
    sheet.setColumnWidth(5, 120);  // Phone
    sheet.setColumnWidth(6, 200);  // Certifications
    sheet.setColumnWidth(7, 200);  // Availability
    sheet.setColumnWidth(8, 250);  // Notes
    
    // Freeze header row
    sheet.setFrozenRows(1);
  }
  
  return sheet;
}

/**
 * Updates the class selector in the Group Lesson Tracker
 * This function refreshes the class dropdown in the dynamic instructor sheet
 * 
 * @returns Success status
 */
function updateClassSelector() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Updating class selector', 'INFO', 'updateClassSelector');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName(DATA_SHEETS.CLASSES);
    
    if (!classesSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Classes sheet not found', 'ERROR', 'updateClassSelector');
      }
      
      SpreadsheetApp.getUi().alert(
        'Missing Data',
        'The Classes sheet is missing. Please initialize the system first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get the data range from the Classes sheet
    const dataRange = classesSheet.getDataRange();
    const classData = dataRange.getValues();
    
    // Skip the header row
    if (classData.length <= 1) {
      SpreadsheetApp.getUi().alert(
        'No Classes Found',
        'There are no classes defined in the system. Please add classes first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Extract class information for the dropdown
    const classOptions = [];
    for (let i = 1; i < classData.length; i++) {
      const classId = classData[i][CLASS_COLUMNS.CLASS_ID];
      const className = classData[i][CLASS_COLUMNS.CLASS_NAME];
      const instructor = classData[i][CLASS_COLUMNS.INSTRUCTOR];
      const day = classData[i][CLASS_COLUMNS.DAY];
      const time = classData[i][CLASS_COLUMNS.TIME];
      
      // Add to options if the class has an ID and name
      if (classId && className) {
        classOptions.push([
          `${className} (${day} ${time}, ${instructor})`
        ]);
      }
    }
    
    // Get the instructor sheet
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    if (!trackerSheet) {
      // If the sheet doesn't exist yet, we'll create it when selected
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Group Lesson Tracker sheet not found, will be created when a class is selected', 'INFO', 'updateClassSelector');
      }
      
      // Create a temporary sheet for the dropdown
      const tempSheet = ss.getSheetByName('ClassSelector') || ss.insertSheet('ClassSelector');
      
      // Clear existing data
      tempSheet.clear();
      
      // Add class options to the sheet
      tempSheet.getRange(1, 1, classOptions.length, 1).setValues(classOptions);
      
      SpreadsheetApp.getUi().alert(
        'Class List Updated',
        'The class list has been refreshed. Please select a class to generate the Group Lesson Tracker.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      return true;
    }
    
    // Update the named range for class selection if it exists
    let classSelectionRange;
    try {
      classSelectionRange = ss.getRangeByName('ClassSelectionList');
    } catch (e) {
      // Named range doesn't exist yet
    }
    
    if (!classSelectionRange) {
      // Create a hidden sheet for the class list if it doesn't exist
      let classListSheet = ss.getSheetByName('ClassList');
      if (!classListSheet) {
        classListSheet = ss.insertSheet('ClassList');
        classListSheet.hideSheet();
      } else {
        classListSheet.clear();
      }
      
      // Add class options to the sheet
      classListSheet.getRange(1, 1, classOptions.length, 1).setValues(classOptions);
      
      // Define the named range
      ss.setNamedRange('ClassSelectionList', classListSheet.getRange(1, 1, classOptions.length, 1));
    } else {
      // Update the existing range
      const sheet = classSelectionRange.getSheet();
      sheet.clear();
      sheet.getRange(1, 1, classOptions.length, 1).setValues(classOptions);
    }
    
    // Update the data validation on the tracker sheet
    try {
      const selectorCell = trackerSheet.getRange('B2'); // Class selector cell
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(ss.getRangeByName('ClassSelectionList'), true)
        .build();
      selectorCell.setDataValidation(rule);
    } catch (e) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Error setting data validation: ${e.message}`, 'ERROR', 'updateClassSelector');
      }
    }
    
    SpreadsheetApp.getUi().alert(
      'Class List Updated',
      'The class list has been refreshed. You can now select a class from the dropdown in cell B2.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'updateClassSelector', 
        'Error updating class selector. Please try again or contact support.');
    } else {
      Logger.log(`Error updating class selector: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Update Failed',
        `Failed to update class selector: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Refreshes the roster data from the source file
 * 
 * @returns Success status
 */
function refreshRosterData() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Refreshing roster data', 'INFO', 'refreshRosterData');
    }
    
    // Get the system configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    if (!config.rosterFolderUrl) {
      SpreadsheetApp.getUi().alert(
        'Missing Configuration',
        'Roster folder URL is not configured. Please update system configuration first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    if (!config.sessionName) {
      SpreadsheetApp.getUi().alert(
        'Missing Configuration',
        'Session name is not configured. Please update system configuration first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Determine folder ID from URL
    let folderId = config.rosterFolderUrl;
    if (folderId.includes('/')) {
      // Extract ID from URL
      const urlPattern = /[-\w]{25,}/;
      const match = folderId.match(urlPattern);
      if (match && match[0]) {
        folderId = match[0];
      }
    }
    
    // Find roster file in the folder
    const folder = DriveApp.getFolderById(folderId);
    const expectedFileName = `YSL ${config.sessionName} Roster`;
    
    // Search for files matching the name
    const files = folder.getFilesByName(expectedFileName);
    
    if (!files.hasNext()) {
      // Try with .xlsx or .csv extensions
      const files2 = folder.getFilesByName(`${expectedFileName}.xlsx`);
      if (!files2.hasNext()) {
        const files3 = folder.getFilesByName(`${expectedFileName}.csv`);
        if (!files3.hasNext()) {
          SpreadsheetApp.getUi().alert(
            'File Not Found',
            `Could not find roster file with name "${expectedFileName}" in the folder. Please check that the file exists and you have permission to access it.`,
            SpreadsheetApp.getUi().ButtonSet.OK
          );
          return false;
        }
      }
    }
    
    // Open the roster file
    const file = files.next();
    const rosterSS = SpreadsheetApp.openById(file.getId());
    const rosterSheet = rosterSS.getSheets()[0]; // Assume first sheet contains roster
    
    // Get roster data
    const rosterData = rosterSheet.getDataRange().getValues();
    
    // Map to our Roster sheet format
    // We'll need to map the columns from the source to our format
    // This is a simplified example; real implementation would have more robust mapping
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(DATA_SHEETS.ROSTER);
    
    if (!targetSheet) {
      // Create Roster sheet if it doesn't exist
      createRosterSheet(ss);
      return refreshRosterData(); // Recursive call now that sheet exists
    }
    
    // Clear existing data (except headers)
    if (targetSheet.getLastRow() > 1) {
      targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
    }
    
    if (rosterData.length <= 1) {
      SpreadsheetApp.getUi().alert(
        'Empty Roster',
        'The roster file does not contain any student data.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Prepare data for import
    const mappedData = [];
    
    // Skip header row
    for (let i = 1; i < rosterData.length; i++) {
      // Basic mapping - this would need to be adjusted based on the actual source format
      const row = rosterData[i];
      if (!row[0]) continue; // Skip empty rows
      
      mappedData.push([
        row[0] || '',                   // Student ID
        row[1] || '',                   // Class ID
        row[2] || '',                   // First Name
        row[3] || '',                   // Last Name
        row[4] || '',                   // Age
        row[5] || '',                   // Level
        row[6] || '',                   // Notes
        row[7] || '',                   // Parent Email
        row[8] || '',                   // Phone
        row[9] || 'Active',             // Status (default to Active)
        row[10] || new Date()           // Assessment Date (default to today)
      ]);
    }
    
    // Write to Roster sheet
    if (mappedData.length > 0) {
      targetSheet.getRange(2, 1, mappedData.length, mappedData[0].length).setValues(mappedData);
    }
    
    // Update the class counts in the Classes sheet
    updateClassCounts();
    
    SpreadsheetApp.getUi().alert(
      'Roster Updated',
      `Successfully imported ${mappedData.length} students from the roster file.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'refreshRosterData', 
        'Error refreshing roster data. Please check the roster file and try again.');
    } else {
      Logger.log(`Error refreshing roster data: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Refresh Failed',
        `Failed to refresh roster data: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Updates the class counts in the Classes sheet based on the Roster
 */
function updateClassCounts() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName(DATA_SHEETS.CLASSES);
    const rosterSheet = ss.getSheetByName(DATA_SHEETS.ROSTER);
    
    if (!classesSheet || !rosterSheet) {
      return; // Can't update counts if sheets don't exist
    }
    
    // Get class data
    const classData = classesSheet.getDataRange().getValues();
    if (classData.length <= 1) return; // No classes defined
    
    // Get roster data
    const rosterData = rosterSheet.getDataRange().getValues();
    if (rosterData.length <= 1) return; // No students defined
    
    // Count students in each class
    const classCounts = {};
    for (let i = 1; i < rosterData.length; i++) {
      const classId = rosterData[i][ROSTER_COLUMNS.CLASS_ID];
      const status = rosterData[i][ROSTER_COLUMNS.STATUS];
      
      // Only count active students
      if (classId && status === 'Active') {
        classCounts[classId] = (classCounts[classId] || 0) + 1;
      }
    }
    
    // Update class sheet with counts
    for (let i = 1; i < classData.length; i++) {
      const classId = classData[i][CLASS_COLUMNS.CLASS_ID];
      if (classId) {
        const count = classCounts[classId] || 0;
        classesSheet.getRange(i + 1, CLASS_COLUMNS.NUM_STUDENTS + 1).setValue(count);
      }
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error updating class counts: ${error.message}`, 'ERROR', 'updateClassCounts');
    } else {
      Logger.log(`Error updating class counts: ${error.message}`);
    }
  }
}

/**
 * Pushes assessment data to the Swimmer Log
 * 
 * @returns Success status
 */
function pushAssessmentsToSwimmerLog() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Pushing assessments to swimmer log', 'INFO', 'pushAssessmentsToSwimmerLog');
    }
    
    // Get the system configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    if (!config.swimmerRecordsUrl) {
      SpreadsheetApp.getUi().alert(
        'Missing Configuration',
        'Swimmer Records URL is not configured. Please update system configuration first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Confirm the operation
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Push Assessments',
      'This will push all assessment data to the Swimmer Records workbook. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return false;
    }
    
    // Open the dynamic instructor sheet to get assessment data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    
    if (!trackerSheet) {
      ui.alert(
        'Missing Data',
        'The Group Lesson Tracker sheet is not found. Please generate it first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Get the class information from the tracker
    const classNameCell = trackerSheet.getRange('B2').getValue();
    if (!classNameCell) {
      ui.alert(
        'Missing Data',
        'No class is selected in the Group Lesson Tracker. Please select a class first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Extract student assessment data
    // This is a simplified example; real implementation would have more robust data extraction
    const dataRange = trackerSheet.getDataRange().getValues();
    
    // Find the student data section (typically starts around row 7)
    let studentStartRow = -1;
    for (let i = 0; i < dataRange.length; i++) {
      if (dataRange[i][0] === 'Student ID' || dataRange[i][0] === 'Student') {
        studentStartRow = i;
        break;
      }
    }
    
    if (studentStartRow === -1) {
      ui.alert(
        'Invalid Format',
        'Could not find student data in the Group Lesson Tracker. The sheet may be formatted incorrectly.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Extract student assessments
    const assessments = [];
    for (let i = studentStartRow + 1; i < dataRange.length; i++) {
      const row = dataRange[i];
      const studentId = row[0];
      const studentName = row[1];
      
      if (!studentId || !studentName) continue; // Skip empty rows
      
      // Collect skill assessments (starting from column index 2)
      const skillAssessments = [];
      for (let j = 2; j < row.length; j++) {
        if (row[j] && row[j] !== '') {
          const skillName = dataRange[studentStartRow][j];
          skillAssessments.push({
            skill: skillName,
            assessment: row[j]
          });
        }
      }
      
      assessments.push({
        studentId: studentId,
        studentName: studentName,
        skills: skillAssessments,
        date: new Date(),
        class: classNameCell
      });
    }
    
    if (assessments.length === 0) {
      ui.alert(
        'No Data',
        'No assessment data found in the Group Lesson Tracker.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Open swimmer records workbook
    let swimmerRecordsId = config.swimmerRecordsUrl;
    if (swimmerRecordsId.includes('/')) {
      // Extract ID from URL
      const urlPattern = /[-\w]{25,}/;
      const match = swimmerRecordsId.match(urlPattern);
      if (match && match[0]) {
        swimmerRecordsId = match[0];
      }
    }
    
    const swimmerSS = SpreadsheetApp.openById(swimmerRecordsId);
    const swimmerLogSheet = swimmerSS.getSheetByName('SwimmerLog') || swimmerSS.getSheets()[0];
    
    // Get existing data to append to
    const swimmerLogData = swimmerLogSheet.getDataRange().getValues();
    const lastRow = swimmerLogSheet.getLastRow();
    
    // Prepare data for import
    const newData = [];
    const sessionName = config.sessionName || '';
    const today = new Date();
    
    assessments.forEach(assessment => {
      assessment.skills.forEach(skill => {
        newData.push([
          assessment.studentId,
          assessment.studentName,
          assessment.class,
          skill.skill,
          skill.assessment,
          today,
          sessionName,
          `Imported from Group Lesson Tracker`
        ]);
      });
    });
    
    // Append to swimmer log
    if (newData.length > 0) {
      swimmerLogSheet.getRange(lastRow + 1, 1, newData.length, newData[0].length).setValues(newData);
    }
    
    ui.alert(
      'Assessment Pushed',
      `Successfully pushed ${newData.length} assessment records to the Swimmer Records workbook.`,
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'pushAssessmentsToSwimmerLog', 
        'Error pushing assessments to swimmer log. Please try again or contact support.');
    } else {
      Logger.log(`Error pushing assessments: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Push Failed',
        `Failed to push assessments: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Pulls assessment criteria from the Swimmer Records workbook
 * 
 * @param swimmerRecordsUrl - Optional URL of the swimmer records workbook
 * @returns Success status
 */
function pullAssessmentCriteria(swimmerRecordsUrl) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Pulling assessment criteria', 'INFO', 'pullAssessmentCriteria');
    }
    
    // Get the system configuration if URL not provided
    let url = swimmerRecordsUrl;
    if (!url) {
      const config = AdministrativeModule.getSystemConfiguration();
      url = config.swimmerRecordsUrl;
    }
    
    if (!url) {
      SpreadsheetApp.getUi().alert(
        'Missing Configuration',
        'Swimmer Records URL is not configured. Please update system configuration first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Extract ID from URL if needed
    let swimmerRecordsId = url;
    if (swimmerRecordsId.includes('/')) {
      const urlPattern = /[-\w]{25,}/;
      const match = swimmerRecordsId.match(urlPattern);
      if (match && match[0]) {
        swimmerRecordsId = match[0];
      }
    }
    
    // Open swimmer records workbook
    const swimmerSS = SpreadsheetApp.openById(swimmerRecordsId);
    
    // Find the assessment criteria sheet
    const criteriaSheet = swimmerSS.getSheetByName('Assessment Criteria') || 
                          swimmerSS.getSheetByName('Criteria') || 
                          swimmerSS.getSheetByName('Skills');
    
    if (!criteriaSheet) {
      // If called from UI (not during initialization)
      if (!swimmerRecordsUrl) {
        SpreadsheetApp.getUi().alert(
          'Sheet Not Found',
          'Assessment Criteria sheet was not found in the Swimmer Records workbook. Please check that the workbook is correctly formatted.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
      return false;
    }
    
    // Get criteria data
    const criteriaData = criteriaSheet.getDataRange().getValues();
    
    // Prepare target sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(DATA_SHEETS.CRITERIA);
    
    if (!targetSheet) {
      // Create AssessmentCriteria sheet if it doesn't exist
      createAssessmentCriteriaSheet(ss);
      return pullAssessmentCriteria(url); // Recursive call now that sheet exists
    }
    
    // Clear existing data (except headers)
    if (targetSheet.getLastRow() > 1) {
      targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
    }
    
    // Skip header row and prepare data for import
    const mappedData = [];
    
    // Skip header row
    for (let i = 1; i < criteriaData.length; i++) {
      const row = criteriaData[i];
      if (!row[0]) continue; // Skip empty rows
      
      // Map source columns to our format - adjust indices based on actual format
      mappedData.push([
        row[0] || '',         // Level
        row[1] || '',         // Category
        row[2] || '',         // Skill
        row[3] || '',         // Description
        row[4] || 'Proficient' // Proficiency Required
      ]);
    }
    
    // Write to criteria sheet
    if (mappedData.length > 0) {
      targetSheet.getRange(2, 1, mappedData.length, mappedData[0].length).setValues(mappedData);
    }
    
    // If called from UI (not during initialization)
    if (!swimmerRecordsUrl) {
      SpreadsheetApp.getUi().alert(
        'Criteria Updated',
        `Successfully imported ${mappedData.length} assessment criteria from the Swimmer Records workbook.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'pullAssessmentCriteria', 
        'Error pulling assessment criteria. Please check the Swimmer Records workbook and try again.');
    } else {
      Logger.log(`Error pulling assessment criteria: ${error.message}`);
      // Only show alert if called from UI
      if (!swimmerRecordsUrl) {
        SpreadsheetApp.getUi().alert(
          'Pull Failed',
          `Failed to pull assessment criteria: ${error.message}`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      }
    }
    return false;
  }
}

/**
 * Imports instructor data from the Session Programs workbook
 * 
 * @param sessionProgramsUrl - URL of the session programs workbook
 * @returns Success status
 */
function importInstructorData(sessionProgramsUrl) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Importing instructor data', 'INFO', 'importInstructorData');
    }
    
    if (!sessionProgramsUrl) {
      const config = AdministrativeModule.getSystemConfiguration();
      sessionProgramsUrl = config.sessionProgramsUrl;
      
      if (!sessionProgramsUrl) {
        SpreadsheetApp.getUi().alert(
          'Missing Configuration',
          'Session Programs URL is not configured. Please update system configuration first.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        return false;
      }
    }
    
    // Extract ID from URL if needed
    let sessionProgramsId = sessionProgramsUrl;
    if (sessionProgramsId.includes('/')) {
      const urlPattern = /[-\w]{25,}/;
      const match = sessionProgramsId.match(urlPattern);
      if (match && match[0]) {
        sessionProgramsId = match[0];
      }
    }
    
    // Open session programs workbook
    const programsSS = SpreadsheetApp.openById(sessionProgramsId);
    
    // Find the instructors sheet
    const instructorsSheet = programsSS.getSheetByName('Instructors') || 
                             programsSS.getSheetByName('Staff') || 
                             programsSS.getSheetByName('Teachers');
    
    if (!instructorsSheet) {
      SpreadsheetApp.getUi().alert(
        'Sheet Not Found',
        'Instructors sheet was not found in the Session Programs workbook. Please check that the workbook is correctly formatted.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get instructor data
    const instructorData = instructorsSheet.getDataRange().getValues();
    
    // Prepare target sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const targetSheet = ss.getSheetByName(DATA_SHEETS.INSTRUCTORS);
    
    if (!targetSheet) {
      // Create Instructors sheet if it doesn't exist
      createInstructorsSheet(ss);
      return importInstructorData(sessionProgramsUrl); // Recursive call now that sheet exists
    }
    
    // Clear existing data (except headers)
    if (targetSheet.getLastRow() > 1) {
      targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, targetSheet.getLastColumn()).clear();
    }
    
    // Skip header row and prepare data for import
    const mappedData = [];
    
    // Skip header row
    for (let i = 1; i < instructorData.length; i++) {
      const row = instructorData[i];
      if (!row[0]) continue; // Skip empty rows
      
      // Map source columns to our format - adjust indices based on actual format
      mappedData.push([
        row[0] || '',         // Instructor ID
        row[1] || '',         // First Name
        row[2] || '',         // Last Name
        row[3] || '',         // Email
        row[4] || '',         // Phone
        row[5] || '',         // Certifications
        row[6] || '',         // Availability
        row[7] || ''          // Notes
      ]);
    }
    
    // Write to instructors sheet
    if (mappedData.length > 0) {
      targetSheet.getRange(2, 1, mappedData.length, mappedData[0].length).setValues(mappedData);
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'importInstructorData', 
        'Error importing instructor data. Please check the Session Programs workbook and try again.');
    } else {
      Logger.log(`Error importing instructor data: ${error.message}`);
      throw error; // Re-throw to caller
    }
    return false;
  }
}

/**
 * Generates a report of assessment criteria for reference
 * 
 * @returns Success status
 */
function reportAssessmentCriteria() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Generating assessment criteria report', 'INFO', 'reportAssessmentCriteria');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const criteriaSheet = ss.getSheetByName(DATA_SHEETS.CRITERIA);
    
    if (!criteriaSheet) {
      SpreadsheetApp.getUi().alert(
        'Missing Data',
        'Assessment criteria data is not available. Please pull assessment criteria first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get criteria data
    const criteriaData = criteriaSheet.getDataRange().getValues();
    
    if (criteriaData.length <= 1) {
      SpreadsheetApp.getUi().alert(
        'No Criteria',
        'There are no assessment criteria defined. Please pull assessment criteria from the Swimmer Records workbook first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Create or get the report sheet
    let reportSheet = ss.getSheetByName('Assessment Criteria Report');
    if (!reportSheet) {
      reportSheet = ss.insertSheet('Assessment Criteria Report');
    } else {
      reportSheet.clear();
    }
    
    // Format the report sheet
    reportSheet.getRange('A1:E1').setValues([criteriaData[0]])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add data
    const reportData = criteriaData.slice(1);
    if (reportData.length > 0) {
      reportSheet.getRange(2, 1, reportData.length, reportData[0].length).setValues(reportData);
    }
    
    // Format and auto-resize columns
    reportSheet.setFrozenRows(1);
    reportSheet.autoResizeColumns(1, 5);
    
    // Activate the report sheet
    reportSheet.activate();
    
    SpreadsheetApp.getUi().alert(
      'Report Generated',
      'Assessment criteria report has been generated.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'reportAssessmentCriteria', 
        'Error generating assessment criteria report. Please try again or contact support.');
    } else {
      Logger.log(`Error generating assessment criteria report: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Report Failed',
        `Failed to generate assessment criteria report: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Diagnoses issues with assessment criteria import
 * 
 * @returns Success status
 */
function diagnoseCriteriaImport() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Diagnosing criteria import issues', 'INFO', 'diagnoseCriteriaImport');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Get the system configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    if (!config.swimmerRecordsUrl) {
      ui.alert(
        'Missing Configuration',
        'Swimmer Records URL is not configured. Please update system configuration first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Extract ID from URL
    let swimmerRecordsId = config.swimmerRecordsUrl;
    if (swimmerRecordsId.includes('/')) {
      const urlPattern = /[-\w]{25,}/;
      const match = swimmerRecordsId.match(urlPattern);
      if (match && match[0]) {
        swimmerRecordsId = match[0];
      }
    }
    
    // Open swimmer records workbook
    let swimmerSS;
    try {
      swimmerSS = SpreadsheetApp.openById(swimmerRecordsId);
    } catch (error) {
      ui.alert(
        'Access Error',
        `Cannot access the Swimmer Records workbook. Error: ${error.message}\n\nPlease check that the URL is correct and you have permission to access the file.`,
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Check for assessment criteria sheet
    const sheets = swimmerSS.getSheets();
    const criteriaSheetNames = ['Assessment Criteria', 'Criteria', 'Skills'];
    let criteriaSheet = null;
    
    for (const name of criteriaSheetNames) {
      criteriaSheet = swimmerSS.getSheetByName(name);
      if (criteriaSheet) break;
    }
    
    if (!criteriaSheet) {
      ui.alert(
        'Sheet Not Found',
        `Could not find a sheet named "Assessment Criteria", "Criteria", or "Skills" in the Swimmer Records workbook.\n\nPlease check that the workbook contains one of these sheets.`,
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Check sheet structure
    const headerRow = criteriaSheet.getRange(1, 1, 1, 5).getValues()[0];
    
    // Expected headers (approximate)
    const expectedHeaders = ['Level', 'Category', 'Skill', 'Description', 'Proficiency'];
    
    // Check if headers match expectations
    const missingHeaders = [];
    for (const header of expectedHeaders) {
      let found = false;
      for (const actual of headerRow) {
        if (actual && actual.toString().toLowerCase().includes(header.toLowerCase())) {
          found = true;
          break;
        }
      }
      if (!found) missingHeaders.push(header);
    }
    
    if (missingHeaders.length > 0) {
      ui.alert(
        'Invalid Format',
        `The Assessment Criteria sheet is missing expected headers: ${missingHeaders.join(', ')}.\n\nThe sheet should have columns for Level, Category, Skill, Description, and Proficiency.`,
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Check data content
    const dataRowCount = criteriaSheet.getLastRow() - 1;
    
    if (dataRowCount <= 0) {
      ui.alert(
        'Empty Sheet',
        'The Assessment Criteria sheet does not contain any data rows.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // All checks passed
    ui.alert(
      'Diagnostic Passed',
      `The Swimmer Records workbook is properly formatted for assessment criteria import.\n\nFound ${dataRowCount} criteria entries in sheet "${criteriaSheet.getName()}".`,
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'diagnoseCriteriaImport', 
        'Error diagnosing criteria import. Please try again or contact support.');
    } else {
      Logger.log(`Error diagnosing criteria import: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Diagnostic Failed',
        `Failed to diagnose criteria import: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

// Global variable export
const DataIntegrationModule = {
  initializeDataStructures,
  updateClassSelector,
  refreshRosterData,
  pushAssessmentsToSwimmerLog,
  pullAssessmentCriteria,
  importInstructorData,
  reportAssessmentCriteria,
  diagnoseCriteriaImport
};