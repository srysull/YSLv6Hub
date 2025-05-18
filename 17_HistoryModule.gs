/**
 * YSL Hub History Module
 * 
 * This module handles the preservation of historical data between sessions.
 * It manages an internal history system that retains student records, assessments,
 * and session information without relying on external storage.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

// Configuration constants
const HISTORY_CONFIG = {
  SHEET_NAME: 'History',
  SECTION_COLORS: {
    HEADER: '#4285F4',
    SELECTOR: '#D9EAD3',
    SESSION_HEADER: '#9FC5E8',
    DATA_HEADER: '#FFE599'
  },
  DATA_TYPES: {
    STUDENT_RECORDS: 'student_records',
    ASSESSMENTS: 'assessments',
    ATTENDANCE: 'attendance',
    COMMUNICATIONS: 'communications'
  },
  MAX_SESSIONS_DISPLAYED: 10 // Number of sessions to show in selector
};

/**
 * Creates or updates the History sheet for storing historical data
 * @return {Sheet} The History sheet
 */
function createHistorySheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(HISTORY_CONFIG.SHEET_NAME);
    
    // If the sheet doesn't exist, create it
    if (!sheet) {
      sheet = ss.insertSheet(HISTORY_CONFIG.SHEET_NAME);
      
      // Set up initial structure
      setupHistorySheetStructure(sheet);
    }
    
    // Return the sheet
    return sheet;
  } catch (error) {
    console.error('Error creating History sheet: ' + error.message);
    // Show a user-friendly error
    SpreadsheetApp.getUi().alert(
      'History Sheet Creation Error',
      'Could not create the History sheet. Please try again or contact support.\n\n' + 
      'Error details: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
}

/**
 * Sets up the basic structure of the History sheet
 * @param {Sheet} sheet - The History sheet
 */
function setupHistorySheetStructure(sheet) {
  // Configure column widths
  sheet.setColumnWidth(1, 150); // Session column
  sheet.setColumnWidth(2, 150); // Student column
  sheet.setColumnWidth(3, 150); // Class column
  sheet.setColumnWidth(4, 500); // Data column
  
  // Add title
  const titleRange = sheet.getRange('A1:D1');
  titleRange.merge()
    .setValue('YSL Hub Session History')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(HISTORY_CONFIG.SECTION_COLORS.HEADER)
    .setFontColor('#FFFFFF')
    .setHorizontalAlignment('center');
  
  // Add description
  const descriptionRange = sheet.getRange('A2:D2');
  descriptionRange.merge()
    .setValue('This sheet stores historical data from previous swim lesson sessions.')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Add session selector
  const selectorRange = sheet.getRange('A3:D3');
  selectorRange.merge()
    .setValue('Select Session: (No sessions archived yet)')
    .setFontWeight('bold')
    .setBackground(HISTORY_CONFIG.SECTION_COLORS.SELECTOR)
    .setHorizontalAlignment('center');
  
  // Add initial placeholder
  const placeholderRange = sheet.getRange('A4:D10');
  placeholderRange.merge()
    .setValue('No historical data has been archived yet.\n\n' +
             'When you transition to a new session, data from the current session will be automatically archived here.')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // Freeze the header rows
  sheet.setFrozenRows(3);
}

/**
 * Archives the current session data to the History sheet
 * @param {string} sessionName - The name of the session being archived
 * @return {boolean} Success status
 */
function archiveCurrentSession(sessionName) {
  try {
    // Get current date for timestamp
    const archiveDate = new Date();
    const archiveTimestamp = archiveDate.toISOString();
    
    // Create or get the History sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let historySheet = ss.getSheetByName(HISTORY_CONFIG.SHEET_NAME);
    
    if (!historySheet) {
      historySheet = createHistorySheet();
    }
    
    // Create a session record
    const sessionData = {
      name: sessionName,
      archiveDate: archiveTimestamp,
      student_records: [],
      assessments: [],
      attendance: [],
      communications: []
    };
    
    // Archive student records from Daxko sheet
    archiveStudentRecords(sessionData);
    
    // Archive assessment data from instructor sheets
    archiveAssessmentData(sessionData);
    
    // Archive attendance data
    archiveAttendanceData(sessionData);
    
    // Archive communication history
    archiveCommunicationHistory(sessionData);
    
    // Store the session data in the History sheet
    storeSessionDataInHistory(historySheet, sessionData);
    
    // Update session selector
    updateSessionSelector(historySheet);
    
    return true;
  } catch (error) {
    console.error('Error archiving session: ' + error.message);
    
    // Log detailed error for troubleshooting
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Session archiving failed: ' + error.message, 'ERROR', 'archiveCurrentSession');
    }
    
    // Show a user-friendly error
    SpreadsheetApp.getUi().alert(
      'Session Archive Error',
      'Could not archive the current session. Please try again or contact support.\n\n' + 
      'Error details: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false;
  }
}

/**
 * Archives student records from the Daxko sheet
 * @param {Object} sessionData - The session data object to populate
 */
function archiveStudentRecords(sessionData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName('Daxko');
    
    if (!daxkoSheet) {
      console.warn('Daxko sheet not found for archiving student records');
      return;
    }
    
    // Get all data from the Daxko sheet
    const data = daxkoSheet.getDataRange().getValues();
    
    // Skip the header row, process student records
    if (data.length > 1) {
      // Identify key columns from the header row
      const headers = data[0];
      const columns = {
        firstName: findColumnIndex(headers, ['First Name', 'FirstName', 'First']),
        lastName: findColumnIndex(headers, ['Last Name', 'LastName', 'Last']),
        email: findColumnIndex(headers, ['Email', 'Email Address', 'ParentEmail']),
        phone: findColumnIndex(headers, ['Phone', 'Phone Number', 'ParentPhone']),
        program: findColumnIndex(headers, ['Program', 'Class', 'Program Name']),
        session: findColumnIndex(headers, ['Session', 'Time', 'Class Time'])
      };
      
      // Process each row as a student record
      for (let i = 1; i < data.length; i++) {
        // Skip rows with missing essential data
        if (!data[i][columns.firstName] || !data[i][columns.lastName]) {
          continue;
        }
        
        // Create a student record object
        const studentRecord = {
          firstName: data[i][columns.firstName],
          lastName: data[i][columns.lastName],
          email: columns.email >= 0 ? data[i][columns.email] : '',
          phone: columns.phone >= 0 ? data[i][columns.phone] : '',
          program: columns.program >= 0 ? data[i][columns.program] : '',
          session: columns.session >= 0 ? data[i][columns.session] : '',
          archivedFrom: 'Daxko',
          archiveDate: new Date().toISOString()
        };
        
        // Add to session data
        sessionData.student_records.push(studentRecord);
      }
    }
    
    console.log(`Archived ${sessionData.student_records.length} student records`);
  } catch (error) {
    console.error('Error archiving student records: ' + error.message);
    // Continue with other archiving operations
  }
}

/**
 * Archives assessment data from instructor sheets
 * @param {Object} sessionData - The session data object to populate
 */
function archiveAssessmentData(sessionData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Look for the instructor sheet
    const instructorSheet = ss.getSheetByName('Instructor Sheet');
    
    if (!instructorSheet) {
      console.log('Instructor Sheet not found for archiving assessment data');
      return;
    }
    
    // Get data from the instructor sheet
    const data = instructorSheet.getDataRange().getValues();
    
    // Extract class information from the first rows
    const classInfo = {
      name: data[1] && data[1][0] ? data[1][0].toString().replace('Class:', '').trim() : 'Unknown Class'
    };
    
    // Find student name columns
    let firstNameCol = -1;
    let lastNameCol = -1;
    let skillStartCol = -1;
    
    // Look for column headers (usually in row 3)
    for (let i = 0; i < data[2].length; i++) {
      const header = data[2][i] ? data[2][i].toString() : '';
      if (header.includes('First Name')) {
        firstNameCol = i;
      } else if (header.includes('Last Name')) {
        lastNameCol = i;
      } else if (header.includes('Day 1') || header.includes('Attendance')) {
        // Skills start after attendance columns
        skillStartCol = i + 1;
        break;
      }
    }
    
    // If we couldn't find the columns, exit
    if (firstNameCol < 0 || lastNameCol < 0) {
      console.log('Could not find name columns in Instructor Sheet');
      return;
    }
    
    // Process each student row (starting from row 4)
    for (let row = 3; row < data.length; row++) {
      // Skip rows without student names
      if (!data[row][firstNameCol] || !data[row][lastNameCol]) {
        continue;
      }
      
      // Create a student assessment record
      const assessmentRecord = {
        firstName: data[row][firstNameCol],
        lastName: data[row][lastNameCol],
        className: classInfo.name,
        skills: {},
        archivedFrom: 'Instructor Sheet',
        archiveDate: new Date().toISOString()
      };
      
      // Collect skill assessments if available
      if (skillStartCol >= 0) {
        for (let col = skillStartCol; col < data[2].length; col++) {
          const skillName = data[2][col] ? data[2][col].toString() : '';
          
          // Skip empty skill names
          if (!skillName) continue;
          
          // Store the skill assessment if it exists
          if (data[row][col]) {
            assessmentRecord.skills[skillName] = data[row][col].toString();
          }
        }
      }
      
      // Add to session data
      sessionData.assessments.push(assessmentRecord);
    }
    
    console.log(`Archived ${sessionData.assessments.length} assessment records`);
  } catch (error) {
    console.error('Error archiving assessment data: ' + error.message);
    // Continue with other archiving operations
  }
}

/**
 * Archives attendance data from instructor sheets
 * @param {Object} sessionData - The session data object to populate
 */
function archiveAttendanceData(sessionData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Look for the instructor sheet
    const instructorSheet = ss.getSheetByName('Instructor Sheet');
    
    if (!instructorSheet) {
      console.log('Instructor Sheet not found for archiving attendance data');
      return;
    }
    
    // Get data from the instructor sheet
    const data = instructorSheet.getDataRange().getValues();
    
    // Extract class information from the first rows
    const classInfo = {
      name: data[1] && data[1][0] ? data[1][0].toString().replace('Class:', '').trim() : 'Unknown Class'
    };
    
    // Find student name columns and attendance columns
    let firstNameCol = -1;
    let lastNameCol = -1;
    let attendanceCols = [];
    
    // Look for column headers (usually in row 3)
    for (let i = 0; i < data[2].length; i++) {
      const header = data[2][i] ? data[2][i].toString() : '';
      if (header.includes('First Name')) {
        firstNameCol = i;
      } else if (header.includes('Last Name')) {
        lastNameCol = i;
      } else if (header.includes('Day ') || header.includes('Attendance')) {
        attendanceCols.push({
          index: i,
          name: header
        });
      }
    }
    
    // If we couldn't find the columns, exit
    if (firstNameCol < 0 || lastNameCol < 0 || attendanceCols.length === 0) {
      console.log('Could not find required columns in Instructor Sheet');
      return;
    }
    
    // Process each student row (starting from row 4)
    for (let row = 3; row < data.length; row++) {
      // Skip rows without student names
      if (!data[row][firstNameCol] || !data[row][lastNameCol]) {
        continue;
      }
      
      // Create an attendance record
      const attendanceRecord = {
        firstName: data[row][firstNameCol],
        lastName: data[row][lastNameCol],
        className: classInfo.name,
        attendance: {},
        archivedFrom: 'Instructor Sheet',
        archiveDate: new Date().toISOString()
      };
      
      // Collect attendance data
      for (const col of attendanceCols) {
        attendanceRecord.attendance[col.name] = data[row][col.index] === true || 
                                             data[row][col.index] === 'TRUE' || 
                                             data[row][col.index] === 'true' || 
                                             data[row][col.index] === 'X' || 
                                             data[row][col.index] === 'x';
      }
      
      // Add to session data
      sessionData.attendance.push(attendanceRecord);
    }
    
    console.log(`Archived ${sessionData.attendance.length} attendance records`);
  } catch (error) {
    console.error('Error archiving attendance data: ' + error.message);
    // Continue with other archiving operations
  }
}

/**
 * Archives communication history
 * @param {Object} sessionData - The session data object to populate
 */
function archiveCommunicationHistory(sessionData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Look for the Communication Log sheet
    const logSheet = ss.getSheetByName('Communication Log');
    
    if (!logSheet) {
      console.log('Communication Log not found for archiving');
      return;
    }
    
    // Get data from the log sheet
    const data = logSheet.getDataRange().getValues();
    
    // Skip if only header row exists
    if (data.length <= 1) {
      console.log('No communications to archive');
      return;
    }
    
    // Process each communication log entry
    for (let i = 1; i < data.length; i++) {
      // Create a communication record object
      const commRecord = {
        date: data[i][0] instanceof Date ? data[i][0].toISOString() : data[i][0].toString(),
        subject: data[i].length > 1 ? data[i][1] : '',
        recipients: data[i].length > 2 ? data[i][2] : '',
        sender: data[i].length > 3 ? data[i][3] : '',
        status: data[i].length > 4 ? data[i][4] : '',
        archivedFrom: 'Communication Log',
        archiveDate: new Date().toISOString()
      };
      
      // Add to session data
      sessionData.communications.push(commRecord);
    }
    
    console.log(`Archived ${sessionData.communications.length} communication records`);
  } catch (error) {
    console.error('Error archiving communication history: ' + error.message);
    // Continue with other archiving operations
  }
}

/**
 * Stores the session data in the History sheet
 * @param {Sheet} historySheet - The History sheet
 * @param {Object} sessionData - The session data to store
 */
function storeSessionDataInHistory(historySheet, sessionData) {
  try {
    // Convert session data to JSON for storage
    const sessionJson = JSON.stringify(sessionData);
    
    // Check if this session already exists in the sheet
    const data = historySheet.getDataRange().getValues();
    let sessionRow = -1;
    
    // Look for the session header row
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'SESSION_DATA' && data[i][1] === sessionData.name) {
        sessionRow = i + 1; // 1-based row index
        break;
      }
    }
    
    // If session exists, update it, otherwise append new data
    if (sessionRow > 0) {
      // Update existing session data
      historySheet.getRange(sessionRow, 3).setValue(sessionJson);
      console.log(`Updated existing session: ${sessionData.name}`);
    } else {
      // Add new session data
      const lastRow = Math.max(historySheet.getLastRow(), 4); // Start after headers
      
      // Add session marker row
      historySheet.getRange(lastRow + 1, 1).setValue('SESSION_DATA');
      historySheet.getRange(lastRow + 1, 2).setValue(sessionData.name);
      historySheet.getRange(lastRow + 1, 3).setValue(sessionJson);
      
      // Add a visual divider
      historySheet.getRange(lastRow + 2, 1, 1, 4).merge()
        .setValue(`Session: ${sessionData.name} - Archived: ${new Date().toLocaleString()}`)
        .setBackground(HISTORY_CONFIG.SECTION_COLORS.SESSION_HEADER)
        .setFontWeight('bold');
      
      console.log(`Added new session: ${sessionData.name}`);
    }
    
    // Update last archived timestamp
    PropertiesService.getDocumentProperties().setProperty('LAST_ARCHIVE_DATE', new Date().toISOString());
    PropertiesService.getDocumentProperties().setProperty('LAST_ARCHIVED_SESSION', sessionData.name);
  } catch (error) {
    console.error('Error storing session data: ' + error.message);
    throw error; // Re-throw this error since it's critical
  }
}

/**
 * Updates the session selector in the History sheet
 * @param {Sheet} historySheet - The History sheet
 */
function updateSessionSelector(historySheet) {
  try {
    // Get all sessions from the sheet
    const data = historySheet.getDataRange().getValues();
    const sessions = [];
    
    // Find session data rows
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'SESSION_DATA' && data[i][1]) {
        sessions.push(data[i][1].toString());
      }
    }
    
    // If no sessions, skip
    if (sessions.length === 0) {
      console.log('No sessions to add to selector');
      return;
    }
    
    // Update the selector text
    historySheet.getRange('A3:D3').merge()
      .setValue('Select Session:')
      .setFontWeight('bold')
      .setBackground(HISTORY_CONFIG.SECTION_COLORS.SELECTOR)
      .setHorizontalAlignment('center');
    
    // Create a dropdown for session selection
    const validationRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(sessions, true)
      .build();
    
    // Add the dropdown to cell E3
    historySheet.getRange('E3').setDataValidation(validationRule);
    
    // Set up an onEdit trigger if not already present
    setupHistorySelectionTrigger();
    
    console.log('Updated session selector with ' + sessions.length + ' sessions');
  } catch (error) {
    console.error('Error updating session selector: ' + error.message);
    // This is non-critical, so continue
  }
}

/**
 * Sets up a trigger to handle history session selection
 */
function setupHistorySelectionTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let hasHistoryTrigger = false;
    
    // Check if trigger already exists
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onHistorySelectionEdit') {
        hasHistoryTrigger = true;
        break;
      }
    }
    
    // Create trigger if it doesn't exist
    if (!hasHistoryTrigger) {
      ScriptApp.newTrigger('onHistorySelectionEdit')
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onEdit()
        .create();
    }
  } catch (error) {
    console.error('Error setting up history selection trigger: ' + error.message);
    // Not critical, so continue
  }
}

/**
 * Handles selection changes in the History sheet
 * @param {Object} e - The onEdit event object
 */
function onHistorySelectionEdit(e) {
  try {
    // Check if edit was in the History sheet
    if (!e || !e.range || e.range.getSheet().getName() !== HISTORY_CONFIG.SHEET_NAME) {
      return;
    }
    
    // Check if the edit was in the session selector (cell E3)
    if (e.range.getRow() === 3 && e.range.getColumn() === 5) {
      const selectedSession = e.value;
      
      if (!selectedSession) {
        return; // No session selected
      }
      
      // Display the selected session data
      displaySessionData(e.range.getSheet(), selectedSession);
    }
  } catch (error) {
    console.error('Error handling history selection: ' + error.message);
    // Don't throw errors in trigger functions
  }
}

/**
 * Displays the selected session data in the History sheet
 * @param {Sheet} sheet - The History sheet
 * @param {string} sessionName - The name of the session to display
 */
function displaySessionData(sheet, sessionName) {
  try {
    // Find the session data
    const data = sheet.getDataRange().getValues();
    let sessionJson = null;
    
    // Look for the session header row
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === 'SESSION_DATA' && data[i][1] === sessionName) {
        sessionJson = data[i][2];
        break;
      }
    }
    
    if (!sessionJson) {
      console.log(`Session data not found for: ${sessionName}`);
      return;
    }
    
    // Parse the session data
    const sessionData = JSON.parse(sessionJson);
    
    // Clear any existing display data
    const lastRow = sheet.getLastRow();
    if (lastRow > 10) { // Only clear if we have data beyond the initial setup
      sheet.getRange(11, 1, lastRow - 10, 4).clear();
    }
    
    // Add a header for the selected session
    sheet.getRange(4, 1, 1, 4).merge()
      .setValue(`Selected Session: ${sessionName} - Archived: ${new Date(sessionData.archiveDate).toLocaleString()}`)
      .setBackground(HISTORY_CONFIG.SECTION_COLORS.SESSION_HEADER)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Start at row 6
    let currentRow = 6;
    
    // Display student records
    if (sessionData.student_records && sessionData.student_records.length > 0) {
      // Add header
      sheet.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`Student Records (${sessionData.student_records.length} students)`)
        .setBackground(HISTORY_CONFIG.SECTION_COLORS.DATA_HEADER)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      currentRow++;
      
      // Add column headers
      sheet.getRange(currentRow, 1).setValue('First Name').setFontWeight('bold');
      sheet.getRange(currentRow, 2).setValue('Last Name').setFontWeight('bold');
      sheet.getRange(currentRow, 3).setValue('Class').setFontWeight('bold');
      sheet.getRange(currentRow, 4).setValue('Contact').setFontWeight('bold');
      
      currentRow++;
      
      // Add student records (limit to first 50 to avoid overwhelming the sheet)
      const recordLimit = Math.min(sessionData.student_records.length, 50);
      for (let i = 0; i < recordLimit; i++) {
        const record = sessionData.student_records[i];
        sheet.getRange(currentRow, 1).setValue(record.firstName);
        sheet.getRange(currentRow, 2).setValue(record.lastName);
        sheet.getRange(currentRow, 3).setValue(record.program);
        sheet.getRange(currentRow, 4).setValue(record.email || record.phone || '');
        currentRow++;
      }
      
      // If there are more records than the limit, add a note
      if (sessionData.student_records.length > recordLimit) {
        sheet.getRange(currentRow, 1, 1, 4).merge()
          .setValue(`... and ${sessionData.student_records.length - recordLimit} more students`)
          .setFontStyle('italic')
          .setHorizontalAlignment('center');
        currentRow++;
      }
      
      // Add space
      currentRow += 2;
    }
    
    // Display assessment data
    if (sessionData.assessments && sessionData.assessments.length > 0) {
      // Add header
      sheet.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`Assessments (${sessionData.assessments.length} students)`)
        .setBackground(HISTORY_CONFIG.SECTION_COLORS.DATA_HEADER)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      currentRow++;
      
      // Add column headers
      sheet.getRange(currentRow, 1).setValue('First Name').setFontWeight('bold');
      sheet.getRange(currentRow, 2).setValue('Last Name').setFontWeight('bold');
      sheet.getRange(currentRow, 3).setValue('Class').setFontWeight('bold');
      sheet.getRange(currentRow, 4).setValue('Skills Summary').setFontWeight('bold');
      
      currentRow++;
      
      // Add assessment records (limit to first 50)
      const recordLimit = Math.min(sessionData.assessments.length, 50);
      for (let i = 0; i < recordLimit; i++) {
        const record = sessionData.assessments[i];
        sheet.getRange(currentRow, 1).setValue(record.firstName);
        sheet.getRange(currentRow, 2).setValue(record.lastName);
        sheet.getRange(currentRow, 3).setValue(record.className);
        
        // Create a skills summary
        let skillsSummary = '';
        const skillKeys = Object.keys(record.skills || {});
        if (skillKeys.length > 0) {
          skillsSummary = `${skillKeys.length} skills assessed`;
        } else {
          skillsSummary = 'No skills recorded';
        }
        
        sheet.getRange(currentRow, 4).setValue(skillsSummary);
        currentRow++;
      }
      
      // If there are more records than the limit, add a note
      if (sessionData.assessments.length > recordLimit) {
        sheet.getRange(currentRow, 1, 1, 4).merge()
          .setValue(`... and ${sessionData.assessments.length - recordLimit} more assessment records`)
          .setFontStyle('italic')
          .setHorizontalAlignment('center');
        currentRow++;
      }
      
      // Add space
      currentRow += 2;
    }
    
    // Display communications
    if (sessionData.communications && sessionData.communications.length > 0) {
      // Add header
      sheet.getRange(currentRow, 1, 1, 4).merge()
        .setValue(`Communications (${sessionData.communications.length} messages)`)
        .setBackground(HISTORY_CONFIG.SECTION_COLORS.DATA_HEADER)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      currentRow++;
      
      // Add column headers
      sheet.getRange(currentRow, 1).setValue('Date').setFontWeight('bold');
      sheet.getRange(currentRow, 2).setValue('Subject').setFontWeight('bold');
      sheet.getRange(currentRow, 3).setValue('Recipients').setFontWeight('bold');
      sheet.getRange(currentRow, 4).setValue('Status').setFontWeight('bold');
      
      currentRow++;
      
      // Add communication records (limit to first 50)
      const recordLimit = Math.min(sessionData.communications.length, 50);
      for (let i = 0; i < recordLimit; i++) {
        const record = sessionData.communications[i];
        
        // Format the date
        let dateValue = record.date;
        try {
          dateValue = new Date(record.date).toLocaleDateString();
        } catch (e) {
          // Use as-is if parsing fails
        }
        
        sheet.getRange(currentRow, 1).setValue(dateValue);
        sheet.getRange(currentRow, 2).setValue(record.subject);
        sheet.getRange(currentRow, 3).setValue(record.recipients);
        sheet.getRange(currentRow, 4).setValue(record.status);
        currentRow++;
      }
      
      // If there are more records than the limit, add a note
      if (sessionData.communications.length > recordLimit) {
        sheet.getRange(currentRow, 1, 1, 4).merge()
          .setValue(`... and ${sessionData.communications.length - recordLimit} more communication records`)
          .setFontStyle('italic')
          .setHorizontalAlignment('center');
        currentRow++;
      }
    }
    
    // Add a note at the bottom about session data storage
    sheet.getRange(currentRow + 2, 1, 1, 4).merge()
      .setValue('Note: Comprehensive session data is stored internally. This is a summary view.')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Scroll to top
    sheet.setActiveRange(sheet.getRange('A1'));
  } catch (error) {
    console.error('Error displaying session data: ' + error.message);
    
    // Show a user-friendly message in the sheet
    try {
      sheet.getRange(4, 1, 1, 4).merge()
        .setValue(`Error displaying session: ${error.message}`)
        .setBackground('#F4CCCC') // Light red
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
    } catch (e) {
      // If even this fails, just log it
      console.error('Could not display error message: ' + e.message);
    }
  }
}

/**
 * Finds the index of a column by header name
 * @param {Array} headers - Array of header values
 * @param {Array|string} possibleNames - Possible names for the column
 * @return {number} The index of the column or -1 if not found
 */
function findColumnIndex(headers, possibleNames) {
  // If possibleNames is a string, convert to array
  const nameOptions = Array.isArray(possibleNames) ? possibleNames : [possibleNames];
  
  // Check each header for a match
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] ? headers[i].toString().toLowerCase() : '';
    
    for (const name of nameOptions) {
      if (header === name.toLowerCase()) {
        return i; // Exact match
      }
    }
  }
  
  // Try partial matches
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i] ? headers[i].toString().toLowerCase() : '';
    
    for (const name of nameOptions) {
      if (header.includes(name.toLowerCase())) {
        return i; // Partial match
      }
    }
  }
  
  return -1; // No match found
}

// Make functions available to other modules
const HistoryModule = {
  createHistorySheet: createHistorySheet,
  archiveCurrentSession: archiveCurrentSession,
  displaySessionData: displaySessionData,
  onHistorySelectionEdit: onHistorySelectionEdit
};