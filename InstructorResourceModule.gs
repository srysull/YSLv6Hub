/**
 * YSL Hub Instructor Resource Module - Modified
 * 
 * This module creates instructor sheets with different layouts for group and private lessons.
 * Group lessons have swimmers across the top and skills down the left.
 * Private lessons have date, student name, instructor, student record, and notes columns.
 * 
 * @author PenBay YMCA
 * @version 1.3
 */

/**
 * Configuration constants for module-specific settings
 */
const INSTRUCTOR_CONFIG = {
  // Format configuration
  SHEET_FORMAT: {
    HEADER_COLOR: '#4285F4',
    SECTION_COLOR: '#E0E0E0',
    SKILLS_COLOR: '#f3f3f3',
    TITLE_FONT_SIZE: 14,
    PAGE_WIDTH: 11,  // Landscape for wider tables
    PAGE_HEIGHT: 8.5
  },
  
  // Data validation options
  VALIDATION: {
    ASSESSMENT_OPTIONS: ['X', '/'],  // X = can perform, / = taught but cannot perform
    ATTENDANCE_OPTIONS: ['PRESENT', 'ABSENT', 'EXCUSED']
  },
  
  // Class types
  CLASS_TYPES: {
    GROUP: 'group',
    PRIVATE: 'private'
  }
};

/**
 * Generates instructor sheets for all selected classes.
 */
function generateInstructorSheets() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Retrieve selected classes information
    const selectedClasses = getSelectedClassesInfo();
    
    // Validate we have selected classes to process
    if (selectedClasses.length === 0) {
      ui.alert(
        'No Classes Selected',
        'Please select at least one class in the Classes sheet by setting the "Select Class" column to "Select".',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Process each selected class
    const results = processSelectedClasses(selectedClasses);
    
    // Show results summary
    showProcessingResults(results);
  } catch (error) {
    // Handle any unexpected errors
    Logger.log(`Error in generateInstructorSheets: ${error.message}`);
    ui.alert('Error', `An unexpected error occurred: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Retrieves information for all classes that have been selected in the Classes sheet.
 * 
 * @return {Array} Array of selected class objects with comprehensive class information
 */
function getSelectedClassesInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesSheet = ss.getSheetByName('Classes');
  
  if (!classesSheet) {
    throw new Error('Classes sheet not found. Please complete system initialization first.');
  }
  
  // Get all data from Classes sheet
  const classData = classesSheet.getDataRange().getValues();
  
  if (classData.length <= 1) {
    return [];
  }
  
  // Find selected classes (marked with 'Select' in column A)
  const selectedClasses = [];
  for (let i = 1; i < classData.length; i++) {
    if (classData[i][0] === 'Select') {
      // Check if this is a private lesson
      const isPrivate = classData[i][1].includes('Private');
      
      // Extract all class information as a structured object
      selectedClasses.push({
        rowIndex: i,
        program: classData[i][1],
        day: classData[i][2],
        // Store the full time string, not just the start time
        time: classData[i][3],
        location: classData[i][4],
        count: classData[i][5],
        instructor: classData[i][6],
        // Derive class type based on program name
        type: isPrivate ? INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE : INSTRUCTOR_CONFIG.CLASS_TYPES.GROUP
      });
    }
  }
  
  return selectedClasses;
}

/**
 * Extracts just the start time from a time string (e.g., "10:00 AM - 10:45 AM" -> "10:00 AM")
 * 
 * @param {string} timeString - The full time string with start and end times
 * @return {string} The start time only
 */
function extractStartTime(timeString) {
  if (!timeString) return '';
  
  // Check if there's a hyphen or dash indicating a time range
  const timeParts = timeString.split(/[-–—]/);
  if (timeParts.length > 1) {
    // Return just the start time and trim any whitespace
    return timeParts[0].trim();
  }
  
  // If no time range found, return the original string
  return timeString;
}

/**
 * Processes each selected class to generate appropriate instructor sheets.
 * 
 * @param {Array} selectedClasses - Array of selected class objects
 * @return {Object} Results object with success and error counts
 */
function processSelectedClasses(selectedClasses) {
  const results = {
    successCount: 0,
    errorCount: 0,
    errors: [] // Store specific errors for logging/diagnostics
  };
  
  selectedClasses.forEach(classInfo => {
    try {
      let sheet;
      
      // Choose the appropriate sheet format based on class type
      if (classInfo.type === INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE) {
        // For private lessons, use the private lesson format
        sheet = createPrivateLessonSheet(classInfo);
      } else {
        // For group classes, use the group class format
        sheet = createGroupClassSheet(classInfo);
      }
      
      if (sheet) {
        results.successCount++;
      }
    } catch (error) {
      results.errorCount++;
      results.errors.push({
        class: `${classInfo.program} (${classInfo.day}, ${classInfo.time})`,
        error: error.message
      });
      
      Logger.log(`Error generating sheet for ${classInfo.program} (${classInfo.day}, ${classInfo.time}): ${error.message}`);
    }
  });
  
  return results;
}

/**
 * Displays processing results to the user.
 * 
 * @param {Object} results - Results object with success and error counts
 */
function showProcessingResults(results) {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Sheet Generation Complete',
    `${results.successCount} instructor sheets generated successfully.\n` +
    (results.errorCount > 0 ? `${results.errorCount} sheets failed. Check logs for details.` : ''),
    ui.ButtonSet.OK
  );
}

/**
 * Creates an instructor sheet based on class info.
 * This is a public interface that routes to the appropriate sheet type.
 * 
 * @param {Object} classInfo - Class information object
 * @return {Sheet} The created instructor sheet
 */
function createInstructorSheet(classInfo) {
  try {
    // Log the start of sheet creation
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Creating instructor sheet for ${classInfo.program} (${classInfo.day}, ${classInfo.time})`, 'INFO', 'createInstructorSheet');
    }
    
    // Determine if this is a private lesson
    const isPrivate = classInfo.program.includes('Private');
    
    // Set class type if not already set
    if (!classInfo.type) {
      classInfo.type = isPrivate ? INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE : INSTRUCTOR_CONFIG.CLASS_TYPES.GROUP;
    }
    
    // Create the appropriate sheet based on class type
    let sheet;
    if (classInfo.type === INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE) {
      sheet = createPrivateLessonSheet(classInfo);
    } else {
      sheet = createGroupClassSheet(classInfo);
    }
    
    // Show confirmation
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Sheet Created',
      `Instructor sheet for ${classInfo.program} (${classInfo.day}, ${classInfo.time}) has been created.`,
      ui.ButtonSet.OK
    );
    
    return sheet;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createInstructorSheet', 
        'Error creating instructor sheet. Please try again or contact support.');
    } else {
      Logger.log(`Error creating instructor sheet: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create instructor sheet: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return null;
  }
}

/**
 * Creates a group class sheet with horizontal layout (swimmers across top, skills down left).
 * This is a placeholder implementation. Replace with your actual implementation.
 * 
 * @param {Object} classInfo - Comprehensive class information object
 * @return {Sheet} The created instructor sheet
 */
function createGroupClassSheet(classInfo) {
  // Log operation
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage(`Creating group class sheet for ${classInfo.program}`, 'INFO', 'createGroupClassSheet');
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `${classInfo.program} ${classInfo.day} ${extractStartTime(classInfo.time)}`;
  
  // Check if sheet already exists and delete it if it does
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // Create a new sheet
  sheet = ss.insertSheet(sheetName);
  
  // Create basic layout
  sheet.getRange('A1:D1').merge()
    .setValue(`Instructor Sheet: ${classInfo.program}`)
    .setFontWeight('bold')
    .setFontSize(INSTRUCTOR_CONFIG.SHEET_FORMAT.TITLE_FONT_SIZE)
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.HEADER_COLOR)
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  // Add class details
  sheet.getRange('A2').setValue('Day:').setFontWeight('bold');
  sheet.getRange('B2').setValue(classInfo.day);
  sheet.getRange('A3').setValue('Time:').setFontWeight('bold');
  sheet.getRange('B3').setValue(classInfo.time);
  sheet.getRange('A4').setValue('Location:').setFontWeight('bold');
  sheet.getRange('B4').setValue(classInfo.location);
  sheet.getRange('A5').setValue('Instructor:').setFontWeight('bold');
  sheet.getRange('B5').setValue(classInfo.instructor);
  
  // Format the sheet
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  
  // You would add more code here to fully implement the group sheet
  // This is a simplified version
  
  return sheet;
}

/**
 * Creates a private lesson sheet.
 * This is a placeholder implementation. Replace with your actual implementation.
 * 
 * @param {Object} classInfo - Comprehensive class information object
 * @return {Sheet} The created instructor sheet
 */
function createPrivateLessonSheet(classInfo) {
  // Log operation
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage(`Creating private lesson sheet for ${classInfo.program}`, 'INFO', 'createPrivateLessonSheet');
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = `${classInfo.program} ${classInfo.day} ${extractStartTime(classInfo.time)}`;
  
  // Check if sheet already exists and delete it if it does
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }
  
  // Create a new sheet
  sheet = ss.insertSheet(sheetName);
  
  // Create basic layout
  sheet.getRange('A1:E1').merge()
    .setValue(`Private Lesson: ${classInfo.program}`)
    .setFontWeight('bold')
    .setFontSize(INSTRUCTOR_CONFIG.SHEET_FORMAT.TITLE_FONT_SIZE)
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.HEADER_COLOR)
    .setFontColor('white')
    .setHorizontalAlignment('center');
  
  // Add class details
  sheet.getRange('A2').setValue('Day:').setFontWeight('bold');
  sheet.getRange('B2').setValue(classInfo.day);
  sheet.getRange('A3').setValue('Time:').setFontWeight('bold');
  sheet.getRange('B3').setValue(classInfo.time);
  sheet.getRange('A4').setValue('Location:').setFontWeight('bold');
  sheet.getRange('B4').setValue(classInfo.location);
  sheet.getRange('A5').setValue('Instructor:').setFontWeight('bold');
  sheet.getRange('B5').setValue(classInfo.instructor);
  
  // Set up column headers for session tracking
  sheet.getRange('A7:E7')
    .setValues([['Date', 'Student Name', 'Instructor', 'Skills Worked On', 'Notes']])
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR);
  
  // Format the sheet
  sheet.setColumnWidth(1, 100);  // Date
  sheet.setColumnWidth(2, 150);  // Student Name
  sheet.setColumnWidth(3, 120);  // Instructor
  sheet.setColumnWidth(4, 200);  // Skills Worked On
  sheet.setColumnWidth(5, 250);  // Notes
  
  // You would add more code here to fully implement the private lesson sheet
  // This is a simplified version
  
  return sheet;
}

// Make functions available to other modules
const InstructorResourceModule = {
  generateInstructorSheets: generateInstructorSheets,
  createInstructorSheet: createInstructorSheet,
  createGroupClassSheet: createGroupClassSheet,
  createPrivateLessonSheet: createPrivateLessonSheet
};