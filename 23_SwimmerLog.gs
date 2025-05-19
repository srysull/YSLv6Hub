/**
 * YSL Hub v2 SwimmerLog Module
 * 
 * This module handles the creation and management of the SwimmerLog sheet,
 * which tracks student attendance and progress across sessions.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-18
 */

/**
 * Creates or updates SwimmerLog and SwimmerSkills sheets based on registration data
 * Uses the session name to create columns for tracking attendance and progress
 * 
 * @param {Object} options - Configuration options
 * @param {string} options.sessionName - The name of the current session
 * @param {boolean} options.overwriteExisting - Whether to overwrite existing student data
 * @param {boolean} options.createSwimmerSkills - Whether to create/update the SwimmerSkills sheet
 * @param {boolean} options.createSwimmerLog - Whether to create/update the SwimmerLog sheet
 * @return {boolean} Success status
 */
function createSwimmerLogs(options) {
  try {
    // Log function entry
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Creating swimmer logs with options: ${JSON.stringify(options)}`, 'INFO', 'createSwimmerLogs');
    } else {
      Logger.log(`Creating swimmer logs with options: ${JSON.stringify(options)}`);
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Validate the session name
    let sessionName = options.sessionName;
    if (!sessionName) {
      // Try to get from dashboard
      if (typeof YSLv6Hub !== 'undefined' && typeof YSLv6Hub.getSessionName === 'function') {
        sessionName = YSLv6Hub.getSessionName();
      }
      
      if (!sessionName) {
        ui.alert('Error', 'Please enter a valid session name.', ui.ButtonSet.OK);
        return false;
      }
    }
    
    // Check if field mappings are complete
    let fieldMappings;
    if (typeof FieldMapping !== 'undefined' && typeof FieldMapping.getFieldMappings === 'function') {
      fieldMappings = FieldMapping.getFieldMappings();
    } else {
      fieldMappings = getFieldMappingsLegacy();
    }
    
    // Verify that required field mappings exist
    const requiredFields = ['firstName', 'lastName', 'dob'];
    const missingFields = requiredFields.filter(field => !fieldMappings[field]);
    
    if (missingFields.length > 0) {
      ui.alert(
        'Missing Field Mappings',
        `The following required fields are not mapped: ${missingFields.join(', ')}. ` +
        'Please complete field mappings before creating swimmer logs.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Get registration data
    const regSheet = ss.getSheetByName('RegInfo');
    if (!regSheet) {
      ui.alert('Error', 'Registration data sheet (RegInfo) not found. Please import registration data first.', ui.ButtonSet.OK);
      return false;
    }
    
    // Update initialization status to "In Progress"
    if (typeof YSLv6Hub !== 'undefined' && typeof YSLv6Hub.updateInitializationStatus === 'function') {
      if (options.createSwimmerLog) {
        YSLv6Hub.updateInitializationStatus('Generate SwimmerLog', 'In Progress');
      }
      if (options.createSwimmerSkills) {
        YSLv6Hub.updateInitializationStatus('Generate SwimmerSkills', 'In Progress');
      }
    }
    
    // Create or update the sheets
    let successCount = 0;
    
    if (options.createSwimmerSkills) {
      const swimmerSkillsSuccess = createOrUpdateSwimmerSkillsSheet(sessionName, fieldMappings, options.overwriteExisting);
      if (swimmerSkillsSuccess) {
        successCount++;
        
        // Update initialization status
        if (typeof YSLv6Hub !== 'undefined' && typeof YSLv6Hub.updateInitializationStatus === 'function') {
          YSLv6Hub.updateInitializationStatus('Generate SwimmerSkills', 'Complete');
        }
      } else {
        // Restore status to pending if failed
        if (typeof YSLv6Hub !== 'undefined' && typeof YSLv6Hub.updateInitializationStatus === 'function') {
          YSLv6Hub.updateInitializationStatus('Generate SwimmerSkills', 'Pending');
        }
      }
    }
    
    if (options.createSwimmerLog) {
      const swimmerLogSuccess = createOrUpdateSwimmerLogSheet(sessionName, fieldMappings, options.overwriteExisting);
      if (swimmerLogSuccess) {
        successCount++;
        
        // Update initialization status
        if (typeof YSLv6Hub !== 'undefined' && typeof YSLv6Hub.updateInitializationStatus === 'function') {
          YSLv6Hub.updateInitializationStatus('Generate SwimmerLog', 'Complete');
        }
      } else {
        // Restore status to pending if failed
        if (typeof YSLv6Hub !== 'undefined' && typeof YSLv6Hub.updateInitializationStatus === 'function') {
          YSLv6Hub.updateInitializationStatus('Generate SwimmerLog', 'Pending');
        }
      }
    }
    
    // Show result message
    if (successCount > 0) {
      ui.alert(
        'Success',
        `Swimmer logs created successfully: ${successCount} sheet(s) updated.`,
        ui.ButtonSet.OK
      );
      return true;
    } else {
      ui.alert(
        'Error',
        'Failed to create swimmer logs. Please check the logs for more information.',
        ui.ButtonSet.OK
      );
      return false;
    }
  } catch (error) {
    // Handle and log errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createSwimmerLogs', 
        'Failed to create swimmer logs. Please check your data and try again.');
    } else {
      Logger.log(`Error creating swimmer logs: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to create swimmer logs: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Creates or updates the SwimmerLog sheet with registration data
 * 
 * @param {string} sessionName - The name of the current session
 * @param {Object} fieldMappings - Mapping of field names to registration columns
 * @param {boolean} overwriteExisting - Whether to overwrite existing student data
 * @return {boolean} Success status
 */
function createOrUpdateSwimmerLogSheet(sessionName, fieldMappings, overwriteExisting) {
  try {
    // Log function entry
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Creating/updating SwimmerLog sheet for session: ${sessionName}`, 'INFO', 'createOrUpdateSwimmerLogSheet');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Define column headers
    const baseHeaders = [
      'First Name',
      'Last Name',
      'DOB',
      'Age',
      'Gender',
      'Program',
      'Class',
      'Notes'
    ];
    
    // Add session columns (current session + placeholders for 8 classes)
    const sessionHeaders = [];
    for (let i = 1; i <= 8; i++) {
      sessionHeaders.push(`[${sessionName}] C${i}`);
    }
    
    // Combine all headers
    const allHeaders = [...baseHeaders, ...sessionHeaders];
    
    // Create or update the SwimmerLog sheet
    let swimmerLogSheet = ss.getSheetByName('SwimmerLog');
    let isNewSheet = false;
    
    if (!swimmerLogSheet) {
      // Create new sheet
      swimmerLogSheet = ss.insertSheet('SwimmerLog');
      isNewSheet = true;
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Created new SwimmerLog sheet', 'INFO', 'createOrUpdateSwimmerLogSheet');
      }
    }
    
    // If new sheet or overwriting existing, set up from scratch
    if (isNewSheet || overwriteExisting) {
      // Clear existing content
      swimmerLogSheet.clear();
      
      // Add headers
      swimmerLogSheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);
      
      // Format header row
      swimmerLogSheet.getRange(1, 1, 1, allHeaders.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      // Freeze header row
      swimmerLogSheet.setFrozenRows(1);
      
      // Set column widths
      swimmerLogSheet.setColumnWidth(1, 120); // First Name
      swimmerLogSheet.setColumnWidth(2, 120); // Last Name
      swimmerLogSheet.setColumnWidth(3, 100); // DOB
      swimmerLogSheet.setColumnWidth(4, 60);  // Age
      swimmerLogSheet.setColumnWidth(5, 80);  // Gender
      swimmerLogSheet.setColumnWidth(6, 120); // Program
      swimmerLogSheet.setColumnWidth(7, 120); // Class
      swimmerLogSheet.setColumnWidth(8, 250); // Notes
      
      // Set column widths for session columns
      for (let i = 0; i < sessionHeaders.length; i++) {
        swimmerLogSheet.setColumnWidth(9 + i, 100); // Session columns
      }
      
      // Create data validation for attendance columns
      const attendanceValues = ['', 'Present', 'Absent', 'Excused', 'Makeup'];
      const attendanceRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(attendanceValues, true)
        .build();
      
      // Apply validation to session columns (starting after the baseHeaders)
      const numRows = 500; // Prepare for many students
      for (let i = 0; i < sessionHeaders.length; i++) {
        swimmerLogSheet.getRange(2, baseHeaders.length + 1 + i, numRows, 1)
          .setDataValidation(attendanceRule);
      }
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Set up SwimmerLog sheet structure', 'INFO', 'createOrUpdateSwimmerLogSheet');
      }
    } else {
      // If not overwriting, check if we need to add new session columns
      const existingHeaders = swimmerLogSheet.getRange(1, 1, 1, swimmerLogSheet.getLastColumn()).getValues()[0];
      
      // Check if session columns already exist
      let allSessionColumnsExist = true;
      for (const sessionHeader of sessionHeaders) {
        if (!existingHeaders.includes(sessionHeader)) {
          allSessionColumnsExist = false;
          break;
        }
      }
      
      if (!allSessionColumnsExist) {
        // Add new session columns
        const lastCol = existingHeaders.length + 1;
        
        // Add session headers
        swimmerLogSheet.getRange(1, lastCol, 1, sessionHeaders.length).setValues([sessionHeaders]);
        
        // Format new headers
        swimmerLogSheet.getRange(1, lastCol, 1, sessionHeaders.length)
          .setFontWeight('bold')
          .setBackground('#f3f3f3');
        
        // Set column widths for new columns
        for (let i = 0; i < sessionHeaders.length; i++) {
          swimmerLogSheet.setColumnWidth(lastCol + i, 100);
        }
        
        // Create data validation for attendance columns
        const attendanceValues = ['', 'Present', 'Absent', 'Excused', 'Makeup'];
        const attendanceRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(attendanceValues, true)
          .build();
        
        // Apply validation to new session columns
        const numRows = Math.max(500, swimmerLogSheet.getLastRow() + 100); // Ensure we cover existing students plus more
        for (let i = 0; i < sessionHeaders.length; i++) {
          swimmerLogSheet.getRange(2, lastCol + i, numRows, 1)
            .setDataValidation(attendanceRule);
        }
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Added ${sessionHeaders.length} new session columns for ${sessionName}`, 'INFO', 'createOrUpdateSwimmerLogSheet');
        }
      }
    }
    
    // Get registration data
    const regSheet = ss.getSheetByName('RegInfo');
    const regData = regSheet.getDataRange().getValues();
    const regHeaders = regData[0];
    
    // Get field mapping indices
    const firstNameColIndex = getColumnIndex(regHeaders, fieldMappings.firstName);
    const lastNameColIndex = getColumnIndex(regHeaders, fieldMappings.lastName);
    const dobColIndex = getColumnIndex(regHeaders, fieldMappings.dob);
    const programColIndex = getColumnIndex(regHeaders, fieldMappings.program);
    const classColIndex = getColumnIndex(regHeaders, fieldMappings.class);
    
    // Prepare student data for import
    const studentData = [];
    
    for (let i = 1; i < regData.length; i++) {
      const regRow = regData[i];
      
      // Skip empty rows
      if (!regRow[firstNameColIndex] && !regRow[lastNameColIndex]) {
        continue;
      }
      
      const firstName = regRow[firstNameColIndex] || '';
      const lastName = regRow[lastNameColIndex] || '';
      const dob = regRow[dobColIndex] || '';
      const program = regRow[programColIndex] || '';
      const className = regRow[classColIndex] || '';
      
      // Calculate age if DOB is available
      let age = '';
      if (dob) {
        try {
          const dobDate = new Date(dob);
          const today = new Date();
          age = Math.floor((today - dobDate) / (365.25 * 24 * 60 * 60 * 1000));
        } catch (e) {
          // Skip age calculation if DOB is invalid
        }
      }
      
      // Create student record
      const studentRecord = [
        firstName,
        lastName,
        dob,
        age,
        '',  // Gender (to be filled in manually)
        program,
        className,
        ''   // Notes
      ];
      
      // Add empty values for session columns
      for (let j = 0; j < sessionHeaders.length; j++) {
        studentRecord.push('');
      }
      
      studentData.push(studentRecord);
    }
    
    if (studentData.length === 0) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('No valid student data found in registration sheet', 'WARNING', 'createOrUpdateSwimmerLogSheet');
      }
      return false;
    }
    
    // Write student data to the SwimmerLog sheet
    if (isNewSheet || overwriteExisting) {
      // Add all student data
      swimmerLogSheet.getRange(2, 1, studentData.length, allHeaders.length).setValues(studentData);
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Added ${studentData.length} students to SwimmerLog sheet`, 'INFO', 'createOrUpdateSwimmerLogSheet');
      }
    } else {
      // Check for existing students and add new ones
      const existingData = swimmerLogSheet.getDataRange().getValues();
      const existingStudents = new Map();
      
      // Create map of existing students (by first name, last name, DOB)
      for (let i = 1; i < existingData.length; i++) {
        const studentKey = `${existingData[i][0]}_${existingData[i][1]}_${existingData[i][2]}`;
        existingStudents.set(studentKey, i + 1); // Row number (1-indexed)
      }
      
      // Count of students added or updated
      let addedCount = 0;
      let updatedCount = 0;
      
      // Process each student from registration data
      for (const student of studentData) {
        const studentKey = `${student[0]}_${student[1]}_${student[2]}`;
        
        if (existingStudents.has(studentKey)) {
          // Student exists, update program and class if needed
          const rowNum = existingStudents.get(studentKey);
          
          // Update program and class columns
          swimmerLogSheet.getRange(rowNum, 6, 1, 2).setValues([[student[5], student[6]]]);
          updatedCount++;
        } else {
          // New student, add to the end
          const newRowNum = swimmerLogSheet.getLastRow() + 1;
          swimmerLogSheet.getRange(newRowNum, 1, 1, student.length).setValues([student]);
          addedCount++;
        }
      }
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Updated SwimmerLog sheet: ${addedCount} students added, ${updatedCount} students updated`, 'INFO', 'createOrUpdateSwimmerLogSheet');
      }
    }
    
    // Apply formatting to the data rows (alternating colors)
    const dataRows = swimmerLogSheet.getLastRow() - 1;
    if (dataRows > 0) {
      for (let i = 0; i < dataRows; i++) {
        if (i % 2 === 1) {
          swimmerLogSheet.getRange(i + 2, 1, 1, allHeaders.length).setBackground('#f5f5f5');
        }
      }
    }
    
    // Activate the SwimmerLog sheet
    swimmerLogSheet.activate();
    
    return true;
  } catch (error) {
    // Handle and log errors
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error creating SwimmerLog sheet: ${error.message}`, 'ERROR', 'createOrUpdateSwimmerLogSheet');
    } else {
      Logger.log(`Error creating SwimmerLog sheet: ${error.message}`);
    }
    return false;
  }
}

/**
 * Creates or updates the SwimmerSkills sheet
 * This is a placeholder implementation - the actual implementation would be more complex
 * 
 * @param {string} sessionName - The name of the current session
 * @param {Object} fieldMappings - Mapping of field names to registration columns
 * @param {boolean} overwriteExisting - Whether to overwrite existing student data
 * @return {boolean} Success status
 */
function createOrUpdateSwimmerSkillsSheet(sessionName, fieldMappings, overwriteExisting) {
  try {
    // Log function entry
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Creating/updating SwimmerSkills sheet for session: ${sessionName}`, 'INFO', 'createOrUpdateSwimmerSkillsSheet');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Define column headers for basic student info
    const baseHeaders = [
      'First Name',
      'Last Name',
      'DOB',
      'Age',
      'Gender',
      'Program',
      'Class',
      'Instructor',
      'Notes'
    ];
    
    // Create or update the SwimmerSkills sheet
    let swimmerSkillsSheet = ss.getSheetByName('SwimmerSkills');
    let isNewSheet = false;
    
    if (!swimmerSkillsSheet) {
      // Create new sheet
      swimmerSkillsSheet = ss.insertSheet('SwimmerSkills');
      isNewSheet = true;
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Created new SwimmerSkills sheet', 'INFO', 'createOrUpdateSwimmerSkillsSheet');
      }
    }
    
    // If new sheet or overwriting existing, set up from scratch
    if (isNewSheet || overwriteExisting) {
      // Clear existing content
      swimmerSkillsSheet.clear();
      
      // Add headers
      swimmerSkillsSheet.getRange(1, 1, 1, baseHeaders.length).setValues([baseHeaders]);
      
      // Format header row
      swimmerSkillsSheet.getRange(1, 1, 1, baseHeaders.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      // Freeze header row
      swimmerSkillsSheet.setFrozenRows(1);
      
      // Set column widths
      swimmerSkillsSheet.setColumnWidth(1, 120); // First Name
      swimmerSkillsSheet.setColumnWidth(2, 120); // Last Name
      swimmerSkillsSheet.setColumnWidth(3, 100); // DOB
      swimmerSkillsSheet.setColumnWidth(4, 60);  // Age
      swimmerSkillsSheet.setColumnWidth(5, 80);  // Gender
      swimmerSkillsSheet.setColumnWidth(6, 120); // Program
      swimmerSkillsSheet.setColumnWidth(7, 120); // Class
      swimmerSkillsSheet.setColumnWidth(8, 120); // Instructor
      swimmerSkillsSheet.setColumnWidth(9, 250); // Notes
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Set up SwimmerSkills sheet structure', 'INFO', 'createOrUpdateSwimmerSkillsSheet');
      }
    }
    
    // Get registration data
    const regSheet = ss.getSheetByName('RegInfo');
    const regData = regSheet.getDataRange().getValues();
    const regHeaders = regData[0];
    
    // Get field mapping indices
    const firstNameColIndex = getColumnIndex(regHeaders, fieldMappings.firstName);
    const lastNameColIndex = getColumnIndex(regHeaders, fieldMappings.lastName);
    const dobColIndex = getColumnIndex(regHeaders, fieldMappings.dob);
    const programColIndex = getColumnIndex(regHeaders, fieldMappings.program);
    const classColIndex = getColumnIndex(regHeaders, fieldMappings.class);
    
    // Prepare student data for import
    const studentData = [];
    
    for (let i = 1; i < regData.length; i++) {
      const regRow = regData[i];
      
      // Skip empty rows
      if (!regRow[firstNameColIndex] && !regRow[lastNameColIndex]) {
        continue;
      }
      
      const firstName = regRow[firstNameColIndex] || '';
      const lastName = regRow[lastNameColIndex] || '';
      const dob = regRow[dobColIndex] || '';
      const program = regRow[programColIndex] || '';
      const className = regRow[classColIndex] || '';
      
      // Calculate age if DOB is available
      let age = '';
      if (dob) {
        try {
          const dobDate = new Date(dob);
          const today = new Date();
          age = Math.floor((today - dobDate) / (365.25 * 24 * 60 * 60 * 1000));
        } catch (e) {
          // Skip age calculation if DOB is invalid
        }
      }
      
      // Create student record
      const studentRecord = [
        firstName,
        lastName,
        dob,
        age,
        '',  // Gender (to be filled in manually)
        program,
        className,
        '',  // Instructor
        ''   // Notes
      ];
      
      studentData.push(studentRecord);
    }
    
    if (studentData.length === 0) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('No valid student data found in registration sheet', 'WARNING', 'createOrUpdateSwimmerSkillsSheet');
      }
      return false;
    }
    
    // Write student data to the SwimmerSkills sheet
    if (isNewSheet || overwriteExisting) {
      // Add all student data
      swimmerSkillsSheet.getRange(2, 1, studentData.length, baseHeaders.length).setValues(studentData);
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Added ${studentData.length} students to SwimmerSkills sheet`, 'INFO', 'createOrUpdateSwimmerSkillsSheet');
      }
    } else {
      // Check for existing students and add new ones
      const existingData = swimmerSkillsSheet.getDataRange().getValues();
      const existingStudents = new Map();
      
      // Create map of existing students (by first name, last name, DOB)
      for (let i = 1; i < existingData.length; i++) {
        const studentKey = `${existingData[i][0]}_${existingData[i][1]}_${existingData[i][2]}`;
        existingStudents.set(studentKey, i + 1); // Row number (1-indexed)
      }
      
      // Count of students added or updated
      let addedCount = 0;
      let updatedCount = 0;
      
      // Process each student from registration data
      for (const student of studentData) {
        const studentKey = `${student[0]}_${student[1]}_${student[2]}`;
        
        if (existingStudents.has(studentKey)) {
          // Student exists, update program and class if needed
          const rowNum = existingStudents.get(studentKey);
          
          // Update program and class columns
          swimmerSkillsSheet.getRange(rowNum, 6, 1, 2).setValues([[student[5], student[6]]]);
          updatedCount++;
        } else {
          // New student, add to the end
          const newRowNum = swimmerSkillsSheet.getLastRow() + 1;
          swimmerSkillsSheet.getRange(newRowNum, 1, 1, student.length).setValues([student]);
          addedCount++;
        }
      }
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Updated SwimmerSkills sheet: ${addedCount} students added, ${updatedCount} students updated`, 'INFO', 'createOrUpdateSwimmerSkillsSheet');
      }
    }
    
    // Apply formatting to the data rows (alternating colors)
    const dataRows = swimmerSkillsSheet.getLastRow() - 1;
    if (dataRows > 0) {
      for (let i = 0; i < dataRows; i++) {
        if (i % 2 === 1) {
          swimmerSkillsSheet.getRange(i + 2, 1, 1, baseHeaders.length).setBackground('#f5f5f5');
        }
      }
    }
    
    return true;
  } catch (error) {
    // Handle and log errors
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error creating SwimmerSkills sheet: ${error.message}`, 'ERROR', 'createOrUpdateSwimmerSkillsSheet');
    } else {
      Logger.log(`Error creating SwimmerSkills sheet: ${error.message}`);
    }
    return false;
  }
}

/**
 * Gets the index of a column in the headers array by its name or first matching item in a list
 * 
 * @param {Array} headers - Array of header strings
 * @param {string|Array} columnName - Column name or array of possible column names
 * @return {number} Column index or -1 if not found
 */
function getColumnIndex(headers, columnName) {
  // If columnName is an array, look for the first match
  if (Array.isArray(columnName)) {
    for (const name of columnName) {
      const index = headers.indexOf(name);
      if (index !== -1) {
        return index;
      }
    }
    return -1;
  }
  
  // Direct match
  return headers.indexOf(columnName);
}

/**
 * Legacy version of getFieldMappings for backward compatibility
 * @return {Object} Field mappings
 */
function getFieldMappingsLegacy() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    const mappingsJson = scriptProperties.getProperty('fieldMappings');
    
    if (mappingsJson) {
      return JSON.parse(mappingsJson);
    }
    
    return {};
  } catch (error) {
    Logger.log(`Error getting field mappings: ${error.message}`);
    return {};
  }
}

/**
 * Updates attendance in the SwimmerLog sheet for a specific student and class
 * 
 * @param {string} firstName - Student's first name
 * @param {string} lastName - Student's last name
 * @param {string} sessionName - The session name
 * @param {number} classNumber - The class number (1-8)
 * @param {string} status - Attendance status ('Present', 'Absent', 'Excused', 'Makeup')
 * @return {boolean} Success status
 */
function updateAttendance(firstName, lastName, sessionName, classNumber, status) {
  try {
    // Log function entry
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Updating attendance for ${firstName} ${lastName}, session ${sessionName}, class ${classNumber} to ${status}`, 'INFO', 'updateAttendance');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swimmerLogSheet = ss.getSheetByName('SwimmerLog');
    
    if (!swimmerLogSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('SwimmerLog sheet not found', 'ERROR', 'updateAttendance');
      }
      return false;
    }
    
    // Get all data
    const data = swimmerLogSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find the student row
    let studentRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === firstName && data[i][1] === lastName) {
        studentRow = i;
        break;
      }
    }
    
    if (studentRow === -1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Student ${firstName} ${lastName} not found in SwimmerLog`, 'ERROR', 'updateAttendance');
      }
      return false;
    }
    
    // Find the session column
    const sessionColumnHeader = `[${sessionName}] C${classNumber}`;
    const sessionCol = headers.indexOf(sessionColumnHeader);
    
    if (sessionCol === -1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Session column ${sessionColumnHeader} not found in SwimmerLog`, 'ERROR', 'updateAttendance');
      }
      return false;
    }
    
    // Update the attendance status
    swimmerLogSheet.getRange(studentRow + 1, sessionCol + 1).setValue(status);
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Successfully updated attendance for ${firstName} ${lastName}`, 'INFO', 'updateAttendance');
    }
    
    return true;
  } catch (error) {
    // Handle and log errors
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error updating attendance: ${error.message}`, 'ERROR', 'updateAttendance');
    } else {
      Logger.log(`Error updating attendance: ${error.message}`);
    }
    return false;
  }
}

/**
 * Gets a list of all students in the SwimmerLog
 * 
 * @return {Array} Array of student objects with firstName, lastName, and rowIndex
 */
function getSwimmerLogStudents() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const swimmerLogSheet = ss.getSheetByName('SwimmerLog');
    
    if (!swimmerLogSheet) {
      return [];
    }
    
    // Get all data
    const data = swimmerLogSheet.getDataRange().getValues();
    
    // Skip header row and extract student info
    const students = [];
    for (let i = 1; i < data.length; i++) {
      const firstName = data[i][0];
      const lastName = data[i][1];
      
      // Skip empty rows
      if (!firstName && !lastName) {
        continue;
      }
      
      students.push({
        firstName: firstName,
        lastName: lastName,
        rowIndex: i + 1 // 1-indexed row number
      });
    }
    
    return students;
  } catch (error) {
    // Handle and log errors
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting SwimmerLog students: ${error.message}`, 'ERROR', 'getSwimmerLogStudents');
    } else {
      Logger.log(`Error getting SwimmerLog students: ${error.message}`);
    }
    return [];
  }
}

// Make functions available to other modules
const SwimmerLog = {
  createSwimmerLogs: createSwimmerLogs,
  createOrUpdateSwimmerLogSheet: createOrUpdateSwimmerLogSheet,
  createOrUpdateSwimmerSkillsSheet: createOrUpdateSwimmerSkillsSheet,
  updateAttendance: updateAttendance,
  getSwimmerLogStudents: getSwimmerLogStudents
};