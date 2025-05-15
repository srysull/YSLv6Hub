/**
 * YSL Hub User Guide Module
 * 
 * This module creates and manages a comprehensive user guide directly within
 * the spreadsheet to help non-technical users understand the system.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

// Configuration constants
const USER_GUIDE_CONFIG = {
  SHEET_NAME: 'YSL Hub User Guide',
  SECTION_COLORS: {
    HEADER: '#4285F4',
    SUBHEADER: '#9FC5E8',
    GETTING_STARTED: '#D9EAD3',
    DAILY_OPERATIONS: '#FFE599',
    TROUBLESHOOTING: '#F4CCCC',
    TRANSITIONS: '#D9D2E9'
  },
  HEADER_TEXT_COLOR: '#FFFFFF'
};

/**
 * Creates or updates the User Guide sheet in the spreadsheet
 * @return {Sheet} The User Guide sheet
 */
function createUserGuideSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(USER_GUIDE_CONFIG.SHEET_NAME);
    
    // Create the sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(USER_GUIDE_CONFIG.SHEET_NAME, 0); // Insert at the beginning
    } else {
      // If it exists, clear its contents
      sheet.clear();
    }
    
    // Configure basic sheet properties
    sheet.setColumnWidth(1, 200); // Section column
    sheet.setColumnWidth(2, 500); // Content column
    sheet.setColumnWidth(3, 300); // Tips column
    
    // Add sheet content
    populateUserGuideContent(sheet);
    
    // Protect the sheet from edits except by owners
    try {
      const protection = sheet.protect().setDescription('User Guide Protection');
      protection.setWarningOnly(true); // Allow edits with warning
    } catch (protectError) {
      console.log('Could not protect User Guide sheet: ' + protectError.message);
      // Continue even if protection fails
    }
    
    // Navigate to the user guide
    sheet.activate();
    ss.setActiveSheet(sheet);
    ss.setActiveSelection('A1');
    
    return sheet;
  } catch (error) {
    console.error('Error creating User Guide: ' + error.message);
    // Show a user-friendly error
    SpreadsheetApp.getUi().alert(
      'User Guide Creation Error',
      'Could not create the User Guide. Please try again or contact support.\n\n' + 
      'Error details: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
}

/**
 * Populates the User Guide sheet with content sections
 * @param {Sheet} sheet - The User Guide sheet
 */
function populateUserGuideContent(sheet) {
  // Add title
  const titleRange = sheet.getRange('A1:C1');
  titleRange.merge()
    .setValue('YSL Hub User Guide')
    .setFontSize(18)
    .setFontWeight('bold')
    .setBackground(USER_GUIDE_CONFIG.SECTION_COLORS.HEADER)
    .setFontColor(USER_GUIDE_CONFIG.HEADER_TEXT_COLOR)
    .setHorizontalAlignment('center');
  
  // Add introduction text
  const introRange = sheet.getRange('A2:C2');
  introRange.merge()
    .setValue('This guide provides instructions for managing your swim lesson program with YSL Hub.')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Current row tracker
  let currentRow = 3;
  
  // Add table of contents
  currentRow = addTableOfContents(sheet, currentRow);
  
  // Add sections
  currentRow = addGettingStartedSection(sheet, currentRow);
  currentRow = addDailyOperationsSection(sheet, currentRow);
  currentRow = addTroubleshootingSection(sheet, currentRow);
  currentRow = addSessionTransitionSection(sheet, currentRow);
  
  // Set up freeze panes to keep headers visible
  sheet.setFrozenRows(3); // Freeze title and intro
}

/**
 * Adds a table of contents to the user guide
 * @param {Sheet} sheet - The User Guide sheet
 * @param {number} startRow - The row to start adding content
 * @return {number} The next available row
 */
function addTableOfContents(sheet, startRow) {
  // Section header
  const headerRange = sheet.getRange(startRow, 1, 1, 3);
  headerRange.merge()
    .setValue('Table of Contents')
    .setFontWeight('bold')
    .setBackground(USER_GUIDE_CONFIG.SECTION_COLORS.SUBHEADER)
    .setHorizontalAlignment('center');
  
  // Add TOC entries
  const tocEntries = [
    '1. Getting Started - System Setup and Initialization',
    '2. Daily Operations - Managing Classes and Students',
    '3. Troubleshooting - Common Issues and Solutions',
    '4. Session Transitions - Moving to a New Session'
  ];
  
  let row = startRow + 1;
  for (const entry of tocEntries) {
    sheet.getRange(row, 1, 1, 3).merge()
      .setValue(entry)
      .setFontWeight('bold');
    row++;
  }
  
  // Add a blank row after TOC
  row++;
  return row;
}

/**
 * Adds the Getting Started section to the user guide
 * @param {Sheet} sheet - The User Guide sheet
 * @param {number} startRow - The row to start adding content
 * @return {number} The next available row
 */
function addGettingStartedSection(sheet, startRow) {
  // Section header
  const headerRange = sheet.getRange(startRow, 1, 1, 3);
  headerRange.merge()
    .setValue('1. Getting Started')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(USER_GUIDE_CONFIG.SECTION_COLORS.GETTING_STARTED)
    .setHorizontalAlignment('center');
  
  let row = startRow + 1;
  
  // System requirements
  sheet.getRange(row, 1).setValue('System Requirements')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'YSL Hub requires:\n' +
    '• Google Sheets access\n' +
    '• Roster data from Daxko or manual entry\n' +
    '• Swim assessment criteria spreadsheet\n' +
    '• Google Drive access for reports'
  );
  sheet.getRange(row, 3).setValue(
    'Tip: Bookmark important spreadsheets for easy access. Keep the URLs handy for initialization.'
  ).setFontStyle('italic');
  
  row++;
  
  // First-time setup
  sheet.getRange(row, 1).setValue('First-Time Setup')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    '1. From the YSL Hub menu, select "System Configuration"\n' +
    '2. Enter your session name (e.g., "Summer 2025")\n' +
    '3. Provide the URLs for your roster folder, assessment criteria, and templates\n' +
    '4. Click "Initialize System"\n' +
    '5. Wait for confirmation that setup is complete'
  );
  sheet.getRange(row, 3).setValue(
    'Warning: Initialization will create several new sheets. Don\'t delete these sheets as they\'re required for system operation.'
  ).setFontStyle('italic');
  
  row++;
  
  // Understanding the interface
  sheet.getRange(row, 1).setValue('Understanding the Menu')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'The YSL Hub menu contains these main sections:\n' +
    '• Class Management - Create and update instructor sheets\n' +
    '• Communications - Send emails and announcements\n' +
    '• System Configuration - Change system settings'
  );
  sheet.getRange(row, 3).setValue(
    'Tip: If you don\'t see the menu, refresh the page or check that you have the necessary permissions.'
  ).setFontStyle('italic');
  
  row++;
  
  // Key sheets explanation
  sheet.getRange(row, 1).setValue('Key Sheets Overview')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'YSL Hub creates several important sheets:\n' +
    '• Classes - Lists all swim classes with checkboxes for selection\n' +
    '• Daxko - Contains student roster information\n' +
    '• Instructor Sheet - Dynamic sheet for tracking student progress\n' +
    '• Communication Hub - Interface for sending emails\n' +
    '• Communication Log - Record of all sent communications'
  );
  sheet.getRange(row, 3).setValue(
    'Note: The "History" sheet stores data from previous sessions and is automatically updated during session transitions.'
  ).setFontStyle('italic');
  
  row += 2; // Add extra space after section
  return row;
}

/**
 * Adds the Daily Operations section to the user guide
 * @param {Sheet} sheet - The User Guide sheet
 * @param {number} startRow - The row to start adding content
 * @return {number} The next available row
 */
function addDailyOperationsSection(sheet, startRow) {
  // Section header
  const headerRange = sheet.getRange(startRow, 1, 1, 3);
  headerRange.merge()
    .setValue('2. Daily Operations')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(USER_GUIDE_CONFIG.SECTION_COLORS.DAILY_OPERATIONS)
    .setHorizontalAlignment('center');
  
  let row = startRow + 1;
  
  // Creating instructor sheets
  sheet.getRange(row, 1).setValue('Creating Class Hubs')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    '1. Select "Class Management > Create Dynamic Class Hub" from the menu\n' +
    '2. When the hub is created, select a class from the dropdown at the top\n' +
    '3. Go back to the menu and select "Class Management > Update with Selected Class"\n' +
    '4. The sheet will populate with student data and assessment fields'
  );
  sheet.getRange(row, 3).setValue(
    'Remember: You must create the hub first, then select a class, then update the hub. This two-step process prevents errors.'
  ).setFontStyle('italic');
  
  row++;
  
  // Using instructor sheets
  sheet.getRange(row, 1).setValue('Using Class Hubs')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'The dynamic Class Hub allows you to:\n' +
    '• Track attendance for each class session\n' +
    '• Record before/after skill assessments\n' +
    '• Add notes about student progress\n' +
    '• Change the selected class using the dropdown\n\n' +
    'Instructors should update the hub after each class session.'
  );
  sheet.getRange(row, 3).setValue(
    'Tip: Use "X" for achieved skills, "/" for in-progress skills, and "?" for skills not yet assessed.'
  ).setFontStyle('italic');
  
  row++;
  
  // Sending communications
  sheet.getRange(row, 1).setValue('Sending Communications')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    '1. Select "Communications > Create Dynamic Communications Hub" from the menu\n' +
    '2. In the new sheet, select recipients using the checkboxes\n' +
    '3. Enter a subject and message body (supports HTML formatting)\n' +
    '4. Check the "Select to send" box at the bottom\n' +
    '5. Go to the menu and select "Communications > Send Selected Communication"'
  );
  sheet.getRange(row, 3).setValue(
    'Available template variables: {{firstName}}, {{lastName}}, {{className}}, {{classTime}}, {{instructor}}. These will be replaced with actual values when sent.'
  ).setFontStyle('italic');
  
  row++;
  
  // Generating reports
  sheet.getRange(row, 1).setValue('Sending Reports')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    '1. Ensure all assessment data is up-to-date in the Class Hubs\n' +
    '2. Select "Communications > Send Mid-Session Reports" or "Send End-Session Reports"\n' +
    '3. Choose which classes to include\n' +
    '4. Confirm the action when prompted\n' +
    '5. Reports will be generated and emailed to parents/guardians'
  );
  sheet.getRange(row, 3).setValue(
    'End-session reports should be sent before transitioning to a new session to ensure all assessment data is recorded.'
  ).setFontStyle('italic');
  
  row++;
  
  // Refreshing data
  sheet.getRange(row, 1).setValue('Refreshing Data')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'If your roster data or class list changes:\n' +
    '1. Select "Class Management > Refresh Class List" to update available classes\n' +
    '2. Select "Class Management > Refresh Roster Data" to update student information\n' +
    '3. Re-create any affected Class Hubs to reflect the new data'
  );
  sheet.getRange(row, 3).setValue(
    'Tip: Always refresh data before creating new hubs to ensure you have the most current information.'
  ).setFontStyle('italic');
  
  row += 2; // Add extra space after section
  return row;
}

/**
 * Adds the Troubleshooting section to the user guide
 * @param {Sheet} sheet - The User Guide sheet
 * @param {number} startRow - The row to start adding content
 * @return {number} The next available row
 */
function addTroubleshootingSection(sheet, startRow) {
  // Section header
  const headerRange = sheet.getRange(startRow, 1, 1, 3);
  headerRange.merge()
    .setValue('3. Troubleshooting')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(USER_GUIDE_CONFIG.SECTION_COLORS.TROUBLESHOOTING)
    .setHorizontalAlignment('center');
  
  let row = startRow + 1;
  
  // Common issues
  sheet.getRange(row, 1).setValue('Missing Student Data')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'If student data is missing in your Class Hub:\n' +
    '1. Check that the student exists in the Daxko sheet\n' +
    '2. Verify the student is assigned to the correct class\n' +
    '3. Select "Class Management > Refresh Roster Data"\n' +
    '4. Re-create the Class Hub and update with the selected class'
  );
  sheet.getRange(row, 3).setValue(
    'If student names appear as "Test Student", this means the system couldn\'t find actual students for the selected class.'
  ).setFontStyle('italic');
  
  row++;
  
  // Menu issues
  sheet.getRange(row, 1).setValue('Missing Menu Items')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'If menu items are missing:\n' +
    '1. Refresh the page (F5 or Cmd+R)\n' +
    '2. Check if the system is properly initialized\n' +
    '3. Go to "System Configuration" to verify initialization status\n' +
    '4. Re-initialize if necessary'
  );
  sheet.getRange(row, 3).setValue(
    'The system must be properly initialized for all menu items to appear.'
  ).setFontStyle('italic');
  
  row++;
  
  // Email issues
  sheet.getRange(row, 1).setValue('Email Sending Failures')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'If emails fail to send:\n' +
    '1. Check the Communication Log for error messages\n' +
    '2. Verify recipient email addresses in the Daxko sheet\n' +
    '3. Ensure you have permission to send emails\n' +
    '4. Try sending to a smaller group first\n' +
    '5. Check your daily email quota (Google limits to 100/day)'
  );
  sheet.getRange(row, 3).setValue(
    'For large groups, consider sending communications over multiple days to avoid hitting quota limits.'
  ).setFontStyle('italic');
  
  row++;
  
  // Data validation issues
  sheet.getRange(row, 1).setValue('Incorrect Skills Data')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'If skills are incorrect or missing:\n' +
    '1. Verify the correct class is selected in the Class Hub\n' +
    '2. Check that the correct stage is associated with the class\n' +
    '3. Look for typos in the class name that might affect stage detection\n' +
    '4. Try adding "S1", "S2", etc. to the class name if stage isn\'t detected'
  );
  sheet.getRange(row, 3).setValue(
    'The system automatically detects swim stages from class names. For example, "Swimming S1 Monday" will show Stage 1 skills.'
  ).setFontStyle('italic');
  
  row++;
  
  // Getting help
  sheet.getRange(row, 1).setValue('Getting Additional Help')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'If you need further assistance:\n' +
    '1. Contact your system administrator\n' +
    '2. Email support at: support@penbaymca.org\n' +
    '3. Check the YMCA Aquatics documentation repository\n' +
    '4. Consult the YSL Hub documentation at: https://docs.ymcasoftware.org/yslhub'
  );
  sheet.getRange(row, 3).setValue(
    'When requesting help, note the specific error messages and what steps you were taking when the issue occurred.'
  ).setFontStyle('italic');
  
  row += 2; // Add extra space after section
  return row;
}

/**
 * Adds the Session Transition section to the user guide
 * @param {Sheet} sheet - The User Guide sheet
 * @param {number} startRow - The row to start adding content
 * @return {number} The next available row
 */
function addSessionTransitionSection(sheet, startRow) {
  // Section header
  const headerRange = sheet.getRange(startRow, 1, 1, 3);
  headerRange.merge()
    .setValue('4. Session Transitions')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground(USER_GUIDE_CONFIG.SECTION_COLORS.TRANSITIONS)
    .setHorizontalAlignment('center');
  
  let row = startRow + 1;
  
  // End of session procedures
  sheet.getRange(row, 1).setValue('End of Session Checklist')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'Before ending a session, complete these tasks:\n' +
    '1. Ensure all assessment data is up-to-date in Class Hubs\n' +
    '2. Send end-session reports to parents\n' +
    '3. Select "Communications > Send End-Session Reports"\n' +
    '4. Archive current session data using "System > Start New Session"\n' +
    '5. Verify all data is correctly archived in the History sheet'
  );
  sheet.getRange(row, 3).setValue(
    'Important: Complete all assessments before transitioning to ensure no data is lost.'
  ).setFontStyle('italic');
  
  row++;
  
  // Starting a new session
  sheet.getRange(row, 1).setValue('Starting a New Session')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    '1. Select "System > Start New Session" from the menu\n' +
    '2. Follow the guided transition wizard\n' +
    '3. Provide the new session name (e.g., "Fall 2025")\n' +
    '4. Enter the URL for the new session\'s roster folder\n' +
    '5. Choose which data to carry forward from the previous session\n' +
    '6. Wait for the transition to complete'
  );
  sheet.getRange(row, 3).setValue(
    'The transition wizard will guide you through each step and automatically archive data from the previous session.'
  ).setFontStyle('italic');
  
  row++;
  
  // Accessing historical data
  sheet.getRange(row, 1).setValue('Accessing Historical Data')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'To view data from previous sessions:\n' +
    '1. Navigate to the "History" sheet\n' +
    '2. Use the session selector at the top to choose a previous session\n' +
    '3. Browse student records, assessments, and communication history\n' +
    '4. Use the search function to find specific students across sessions'
  );
  sheet.getRange(row, 3).setValue(
    'Historical data is stored within the system for easy reference without needing external spreadsheets.'
  ).setFontStyle('italic');
  
  row++;
  
  // Transition troubleshooting
  sheet.getRange(row, 1).setValue('Transition Troubleshooting')
    .setFontWeight('bold');
  sheet.getRange(row, 2).setValue(
    'If you encounter issues during transition:\n' +
    '1. Check that all required URLs are correct\n' +
    '2. Verify you have edit access to all referenced spreadsheets\n' +
    '3. Ensure you\'ve completed all end-of-session tasks\n' +
    '4. If transition fails, try again with "System > Resume Transition"\n' +
    '5. For persistent issues, contact technical support'
  );
  sheet.getRange(row, 3).setValue(
    'The system maintains a backup of your previous session configuration until the transition is successfully completed.'
  ).setFontStyle('italic');
  
  row += 2; // Add extra space after section
  
  // Add footer with version info
  const footerRange = sheet.getRange(row, 1, 1, 3);
  footerRange.merge()
    .setValue('YSL Hub User Guide | Updated: ' + new Date().toLocaleDateString())
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  return row + 1;
}

// Make functions available to other modules
const UserGuide = {
  createUserGuideSheet: createUserGuideSheet
};