/**
 * YSL Hub v2 User Guide Module
 * 
 * This module provides functionality for creating and managing a comprehensive
 * user guide for the YSL Hub system. It generates documentation about the
 * system's features and how to use them.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

// User guide sections
const GUIDE_SECTIONS = {
  INTRODUCTION: 'Introduction',
  GETTING_STARTED: 'Getting Started',
  MENU_REFERENCE: 'Menu Reference',
  CLASS_MANAGEMENT: 'Class Management',
  STUDENT_TRACKING: 'Student Tracking',
  ASSESSMENTS: 'Assessments',
  REPORTING: 'Reporting',
  COMMUNICATIONS: 'Communications',
  SYSTEM_ADMIN: 'System Administration',
  FAQ: 'Frequently Asked Questions',
  TROUBLESHOOTING: 'Troubleshooting'
};

/**
 * Creates a user guide sheet with comprehensive documentation
 * 
 * @returns Success status
 */
function createUserGuideSheet() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Creating user guide sheet', 'INFO', 'createUserGuideSheet');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let guideSheet = ss.getSheetByName('User Guide');
    
    // Create the sheet if it doesn't exist
    if (!guideSheet) {
      guideSheet = ss.insertSheet('User Guide');
    } else {
      // Clear existing content
      guideSheet.clear();
    }
    
    // Set up title
    guideSheet.getRange('A1:H1').merge()
      .setValue('YSL Hub System - User Guide')
      .setFontWeight('bold')
      .setFontSize(16)
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Add system information
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || 'Current Session';
    
    let versionInfo = 'Version 2.0.0 (2025-04-27)';
    if (VersionControl && typeof VersionControl.getVersionInfo === 'function') {
      const vi = VersionControl.getVersionInfo();
      versionInfo = `Version ${vi.currentVersion} (${vi.releaseDate})`;
    }
    
    guideSheet.getRange('A2:H2').merge()
      .setValue(`${sessionName} - ${versionInfo}`)
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Add table of contents
    guideSheet.getRange('A4:H4').merge()
      .setValue('Table of Contents')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    let currentRow = 5;
    Object.values(GUIDE_SECTIONS).forEach((section, index) => {
      guideSheet.getRange(currentRow, 1, 1, 8).merge()
        .setValue(`${index + 1}. ${section}`)
        .setFontStyle('italic');
      currentRow++;
    });
    
    currentRow++; // Add spacing
    
    // Add each section content
    Object.entries(GUIDE_SECTIONS).forEach(([key, section]) => {
      // Add section header
      guideSheet.getRange(currentRow, 1, 1, 8).merge()
        .setValue(section)
        .setFontWeight('bold')
        .setBackground('#4285F4')
        .setFontColor('white');
      currentRow++;
      
      // Add section content
      const content = getSectionContent(key, config);
      
      for (const paragraph of content) {
        // If it's a sub-header
        if (paragraph.type === 'header') {
          guideSheet.getRange(currentRow, 1, 1, 8).merge()
            .setValue(paragraph.text)
            .setFontWeight('bold')
            .setBackground('#E1E1E1');
        } 
        // If it's a regular paragraph
        else if (paragraph.type === 'paragraph') {
          guideSheet.getRange(currentRow, 1, 1, 8).merge()
            .setValue(paragraph.text)
            .setWrap(true);
        }
        // If it's a note or tip
        else if (paragraph.type === 'note') {
          guideSheet.getRange(currentRow, 1, 1, 8).merge()
            .setValue(`Note: ${paragraph.text}`)
            .setFontStyle('italic')
            .setBackground('#F0F8FF')
            .setWrap(true);
        }
        // If it's a list
        else if (paragraph.type === 'list') {
          for (const item of paragraph.items) {
            guideSheet.getRange(currentRow, 2, 1, 7).merge()
              .setValue(`• ${item}`)
              .setWrap(true);
            currentRow++;
          }
          // Decrement row counter since we'll increment at the end of the loop
          currentRow--;
        }
        
        currentRow++;
      }
      
      currentRow++; // Add spacing between sections
    });
    
    // Format sheet
    guideSheet.setColumnWidth(1, 30);
    for (let i = 2; i <= 7; i++) {
      guideSheet.setColumnWidth(i, 120);
    }
    guideSheet.setColumnWidth(8, 30);
    
    // Freeze header row
    guideSheet.setFrozenRows(1);
    
    // Ensure the guide sheet is visible and active
    guideSheet.showSheet();
    guideSheet.activate();
    
    SpreadsheetApp.getUi().alert(
      'User Guide Created',
      'The User Guide has been created. You can now browse the documentation for the YSL Hub system.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createUserGuideSheet', 
        'Error creating user guide. Please try again or contact support.');
    } else {
      Logger.log(`Error creating user guide: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create user guide: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Gets content for a specific user guide section
 * 
 * @param sectionKey - The section key
 * @param config - The system configuration
 * @returns Array of content paragraphs
 */
function getSectionContent(sectionKey, config) {
  const content = [];
  const sessionName = config.sessionName || 'Current Session';
  
  switch (sectionKey) {
    case 'INTRODUCTION':
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub system is a comprehensive management tool for swim lessons, ' +
              'providing class management, assessment tracking, report generation, ' +
              'and communication capabilities for the YMCA Aquatics Department.'
      });
      content.push({
        type: 'paragraph',
        text: 'This guide provides detailed information about how to use the system effectively. ' +
              'Please review each section to understand the full capabilities of the YSL Hub.'
      });
      content.push({
        type: 'header',
        text: 'System Purpose'
      });
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub system is designed to streamline and enhance the management of swim lessons by:'
      });
      content.push({
        type: 'list',
        items: [
          'Centralizing class and student data',
          'Standardizing assessment processes',
          'Automating report generation',
          'Facilitating communication with parents',
          'Providing historical tracking of student progress'
        ]
      });
      break;
      
    case 'GETTING_STARTED':
      content.push({
        type: 'header',
        text: 'System Setup'
      });
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub system has been configured for ' + sessionName + '. ' +
              'The system is ready to use with the following components:'
      });
      content.push({
        type: 'list',
        items: [
          'Menu System: Access all functionality from the YSL v6 Hub menu',
          'Data Sheets: Automated creation of necessary data sheets',
          'Sync Functions: Integration with swimmer records',
          'Report Templates: Pre-configured templates for progress reports'
        ]
      });
      content.push({
        type: 'header',
        text: 'Initial Steps'
      });
      content.push({
        type: 'paragraph',
        text: 'To begin using the system, follow these steps:'
      });
      content.push({
        type: 'list',
        items: [
          '1. Generate the Group Lesson Tracker using the menu option',
          '2. Refresh the Class List to populate available classes',
          '3. Select a class to work with',
          '4. Enter assessment data for students'
        ]
      });
      content.push({
        type: 'note',
        text: 'If classes do not appear in the dropdown, you may need to import roster data or refresh class information using the menu options.'
      });
      break;
      
    case 'MENU_REFERENCE':
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub menu is the primary interface for accessing system functionality. ' +
              'Below is a reference for all menu options:'
      });
      content.push({
        type: 'header',
        text: 'Main Menu Options'
      });
      content.push({
        type: 'list',
        items: [
          'Generate Group Lesson Tracker: Creates the main tracking sheet for class assessments',
          'SYNC STUDENT DATA: Synchronizes student data between sheets',
          'Refresh Class List: Updates the class selection dropdown',
          'Refresh Roster Data: Imports the latest student enrollment data'
        ]
      });
      content.push({
        type: 'header',
        text: 'Communications Menu'
      });
      content.push({
        type: 'list',
        items: [
          'Create Communications Hub: Creates the email communication center',
          'Create Communication Log: Sets up tracking for sent communications',
          'Send Selected Communication: Sends emails based on Communications Hub settings',
          'Send Mid-Session Reports: Generates and optionally emails mid-session progress reports',
          'Send End-Session Reports: Generates and optionally emails end-session assessments',
          'Send Welcome Emails: Sends class welcome emails to parents'
        ]
      });
      content.push({
        type: 'header',
        text: 'System Menu'
      });
      content.push({
        type: 'list',
        items: [
          'Create User Guide: Generates this documentation',
          'View History: Shows system update history',
          'Start New Session: Begins a new session transition process',
          'Resume Session Transition: Continues an interrupted session transition',
          'System Configuration: Update system settings',
          'Fix Swimmer Records Access: Resolves access issues with swimmer records',
          'Apply Configuration Changes: Applies changes made in the configuration sheet',
          'Show Logs: Display system logs for troubleshooting'
        ]
      });
      break;
      
    case 'CLASS_MANAGEMENT':
      content.push({
        type: 'header',
        text: 'Group Lesson Tracker'
      });
      content.push({
        type: 'paragraph',
        text: 'The Group Lesson Tracker is the primary tool for managing class information and assessments. ' +
              'It provides a centralized view of all students in a class and their progress on each skill.'
      });
      content.push({
        type: 'list',
        items: [
          '1. Generate the tracker using "Generate Group Lesson Tracker" from the menu',
          '2. Select a class from the dropdown in cell B2',
          '3. The tracker will populate with student information and assessment criteria',
          '4. Enter assessment values in the grid (P = Proficient, D = Developing, etc.)'
        ]
      });
      content.push({
        type: 'header',
        text: 'Class Selection and Management'
      });
      content.push({
        type: 'paragraph',
        text: 'The system draws class information from the Classes sheet, which is synchronized with your roster data. ' +
              'To update class information:'
      });
      content.push({
        type: 'list',
        items: [
          'Use "Refresh Class List" to update the class selection dropdown',
          'Use "Refresh Roster Data" to import the latest enrollment information',
          'The Classes sheet contains detailed information about each class including instructor, schedule, and enrollment'
        ]
      });
      content.push({
        type: 'note',
        text: 'Changes to class schedules, instructors, or other details should be made in the Classes sheet directly.'
      });
      break;
      
    case 'STUDENT_TRACKING':
      content.push({
        type: 'header',
        text: 'Student Data Management'
      });
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub system maintains student information in the Roster sheet. This data is synchronized with ' +
              'the Group Lesson Tracker and can be updated as needed.'
      });
      content.push({
        type: 'paragraph',
        text: 'Student information includes:'
      });
      content.push({
        type: 'list',
        items: [
          'Student ID: Unique identifier for each student',
          'Class assignment: Which class the student is enrolled in',
          'Name and contact information: Student details and parent contact info',
          'Level and age: Student\'s swim level and age',
          'Assessment data: Progress on skills and completion status'
        ]
      });
      content.push({
        type: 'header',
        text: 'Synchronization'
      });
      content.push({
        type: 'paragraph',
        text: 'To ensure the system has up-to-date student information, use the following functions:'
      });
      content.push({
        type: 'list',
        items: [
          'SYNC STUDENT DATA: Updates data between Group Lesson Tracker and SwimmerSkills sheets',
          'Refresh Roster Data: Imports the latest student enrollment from the roster file',
          'Push Assessments to Swimmer Log: Updates the swimmer records with assessment data'
        ]
      });
      content.push({
        type: 'note',
        text: 'Always synchronize data before generating reports to ensure the most current information is used.'
      });
      break;
      
    case 'ASSESSMENTS':
      content.push({
        type: 'header',
        text: 'Assessment Process'
      });
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub system uses standardized assessment criteria for each swim level. These criteria are imported ' +
              'from the Swimmer Records workbook and presented in the Group Lesson Tracker.'
      });
      content.push({
        type: 'paragraph',
        text: 'To conduct assessments:'
      });
      content.push({
        type: 'list',
        items: [
          '1. Open the Group Lesson Tracker for the class',
          '2. For each student, evaluate performance on each skill',
          '3. Enter assessment values using the standard notation:',
          '   - P or Proficient: Student has mastered the skill',
          '   - D or Developing: Student is making progress but needs more work',
          '   - N or Not Yet: Student has not demonstrated the skill',
          '   - Empty: Skill has not been assessed'
        ]
      });
      content.push({
        type: 'header',
        text: 'Assessment Criteria'
      });
      content.push({
        type: 'paragraph',
        text: 'Assessment criteria are organized by swimming level and skill category. The system maintains these ' +
              'criteria in the AssessmentCriteria sheet.'
      });
      content.push({
        type: 'paragraph',
        text: 'To update assessment criteria:'
      });
      content.push({
        type: 'list',
        items: [
          'Use "Pull Assessment Criteria" to import the latest criteria from Swimmer Records',
          'Use "Report Assessment Criteria" to generate a summary sheet of all criteria'
        ]
      });
      content.push({
        type: 'note',
        text: 'If you have trouble importing criteria, use "Diagnose Criteria Import" to identify any issues.'
      });
      break;
      
    case 'REPORTING':
      content.push({
        type: 'header',
        text: 'Progress Reports'
      });
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub system can generate both mid-session progress reports and end-session assessment reports. ' +
              'These reports provide parents with detailed information about their child\'s progress.'
      });
      content.push({
        type: 'paragraph',
        text: 'To generate reports:'
      });
      content.push({
        type: 'list',
        items: [
          '1. Use "Send Mid-Session Reports" or "Send End-Session Reports" from the Communications menu',
          '2. Select the classes to include in the reports',
          '3. Specify a folder name for the reports (or use the default)',
          '4. The system will generate individual reports for each student',
          '5. Optionally send the reports via email to parents'
        ]
      });
      content.push({
        type: 'header',
        text: 'Report Distribution'
      });
      content.push({
        type: 'paragraph',
        text: 'Reports can be distributed in several ways:'
      });
      content.push({
        type: 'list',
        items: [
          'Email: Send reports directly to parents via email',
          'Google Drive sharing: Share the report folder with parents',
          'Print: Print physical copies of reports for distribution',
          'Summary Sheet: Use the generated summary sheet to manage and track reports'
        ]
      });
      content.push({
        type: 'note',
        text: 'Reports are generated as Google Documents and saved in a Google Drive folder. You will need appropriate ' +
              'permissions to access and share these documents.'
      });
      break;
      
    case 'COMMUNICATIONS':
      content.push({
        type: 'header',
        text: 'Communications Hub'
      });
      content.push({
        type: 'paragraph',
        text: 'The Communications Hub is a central tool for creating and sending emails to parents. It provides ' +
              'templates for common communications and tracks message history.'
      });
      content.push({
        type: 'paragraph',
        text: 'To use the Communications Hub:'
      });
      content.push({
        type: 'list',
        items: [
          '1. Create the hub using "Create Communications Hub" from the menu',
          '2. Select a communication type (Class Email, Welcome Email, etc.)',
          '3. Choose a class and template',
          '4. Customize the subject and message as needed',
          '5. Send the communication using "Send Selected Communication"'
        ]
      });
      content.push({
        type: 'header',
        text: 'Communication Types'
      });
      content.push({
        type: 'paragraph',
        text: 'The system supports several types of communications:'
      });
      content.push({
        type: 'list',
        items: [
          'Class Email: General communication to all parents in a class',
          'Welcome Email: Initial welcome for new students',
          'Announcement: Important news or updates for a class',
          'Ready Notice: Notification that a student is ready for the next level',
          'Custom: User-defined communication for special purposes'
        ]
      });
      content.push({
        type: 'header',
        text: 'Communication Log'
      });
      content.push({
        type: 'paragraph',
        text: 'All communications are tracked in the Communication Log, which provides a history of messages sent to parents.'
      });
      content.push({
        type: 'paragraph',
        text: 'The log includes:'
      });
      content.push({
        type: 'list',
        items: [
          'Date and time: When the communication was sent',
          'Type: The type of communication',
          'Recipients: How many people received the message',
          'Subject: The email subject line',
          'Status: Whether the sending was successful',
          'Notes: Additional details about the communication'
        ]
      });
      content.push({
        type: 'note',
        text: 'The Communication Log is a valuable record for administrative purposes and parent communications tracking.'
      });
      break;
      
    case 'SYSTEM_ADMIN':
      content.push({
        type: 'header',
        text: 'System Configuration'
      });
      content.push({
        type: 'paragraph',
        text: 'The YSL Hub system configuration is managed through the Assumptions sheet and script properties. ' +
              'Administrators can modify these settings as needed.'
      });
      content.push({
        type: 'paragraph',
        text: 'Key configuration parameters include:'
      });
      content.push({
        type: 'list',
        items: [
          'Session Name: The current session identifier',
          'Roster Folder URL: Location of class roster files',
          'Report Template URL: Location of report templates',
          'Swimmer Records URL: Link to the swimmer records workbook',
          'Parent Handbook URL: Link to the parent handbook document',
          'Session Programs URL: Link to the session programs workbook'
        ]
      });
      content.push({
        type: 'header',
        text: 'Session Transition'
      });
      content.push({
        type: 'paragraph',
        text: 'At the end of each session, the system can be transitioned to a new session. This process preserves ' +
              'historical data while setting up the system for a new group of classes.'
      });
      content.push({
        type: 'paragraph',
        text: 'The session transition process includes:'
      });
      content.push({
        type: 'list',
        items: [
          '1. Archiving current session data',
          '2. Updating configuration for the new session',
          '3. Importing new roster information',
          '4. Resetting assessment tracking',
          '5. Creating new class structures'
        ]
      });
      content.push({
        type: 'note',
        text: 'Always back up your current workbook before starting a session transition.'
      });
      break;
      
    case 'FAQ':
      content.push({
        type: 'header',
        text: 'Common Questions'
      });
      content.push({
        type: 'paragraph',
        text: 'Q: How do I add a new student to a class?'
      });
      content.push({
        type: 'paragraph',
        text: 'A: Add the student to the Roster sheet with the appropriate class ID, then use "Refresh Class List" and "SYNC STUDENT DATA" to update the system.'
      });
      content.push({
        type: 'paragraph',
        text: 'Q: How do I change a student\'s class assignment?'
      });
      content.push({
        type: 'paragraph',
        text: 'A: Update the Class ID in the Roster sheet for the student, then synchronize the data.'
      });
      content.push({
        type: 'paragraph',
        text: 'Q: What if the menu options aren\'t showing correctly?'
      });
      content.push({
        type: 'paragraph',
        text: 'A: Use "Fix Menu" from the YSL v6 Hub menu to restore proper menu functionality.'
      });
      content.push({
        type: 'paragraph',
        text: 'Q: How do I update assessment criteria for a new program structure?'
      });
      content.push({
        type: 'paragraph',
        text: 'A: Update the criteria in the Swimmer Records workbook, then use "Pull Assessment Criteria" to import them.'
      });
      content.push({
        type: 'paragraph',
        text: 'Q: Can I customize the report templates?'
      });
      content.push({
        type: 'paragraph',
        text: 'A: Yes, edit the templates in Google Drive, then update the Report Template URL in system configuration.'
      });
      break;
      
    case 'TROUBLESHOOTING':
      content.push({
        type: 'header',
        text: 'Common Issues'
      });
      content.push({
        type: 'paragraph',
        text: 'Issue: Menu options are missing or don\'t work'
      });
      content.push({
        type: 'paragraph',
        text: 'Solution: Use "Fix Menu" or "Repair Menu" from the YSL v6 Hub menu to reset the menu system. If that doesn\'t work, try refreshing the page or reopening the spreadsheet.'
      });
      content.push({
        type: 'paragraph',
        text: 'Issue: Cannot access Swimmer Records'
      });
      content.push({
        type: 'paragraph',
        text: 'Solution: Use "Fix Swimmer Records Access" to update the URL and permissions for the Swimmer Records workbook.'
      });
      content.push({
        type: 'paragraph',
        text: 'Issue: Student data is not synchronized correctly'
      });
      content.push({
        type: 'paragraph',
        text: 'Solution: Use "SYNC STUDENT DATA" to force a synchronization between sheets. If issues persist, check for data format problems in the Roster sheet.'
      });
      content.push({
        type: 'header',
        text: 'Error Logs'
      });
      content.push({
        type: 'paragraph',
        text: 'The system maintains detailed logs of operations and errors. To access these logs:'
      });
      content.push({
        type: 'list',
        items: [
          '1. Use "Show Logs" from the System menu',
          '2. Review the SystemLog sheet for error messages and warnings',
          '3. Use the information to identify and resolve issues'
        ]
      });
      content.push({
        type: 'paragraph',
        text: 'For persistent issues, contact the system administrator with the following information:'
      });
      content.push({
        type: 'list',
        items: [
          'Error messages from the log',
          'Steps to reproduce the issue',
          'Screenshot of the problem (if applicable)',
          'Description of what you were trying to do'
        ]
      });
      break;
      
    default:
      content.push({
        type: 'paragraph',
        text: 'No content available for this section.'
      });
  }
  
  return content;
}

// Global variable export
const UserGuide = {
  createUserGuideSheet
};