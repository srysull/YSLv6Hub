/**
 * YSL Hub v2 Communication Module
 * 
 * This module handles all communication functions for the YSL Hub system,
 * including email communications, class announcements, and welcome emails.
 * It provides utilities for creating communication templates and sending
 * communications to parents and students.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

// Communication types
const COMMUNICATION_TYPES = {
  CLASS_EMAIL: 'Class Email',
  WELCOME_EMAIL: 'Welcome Email',
  ANNOUNCEMENT: 'Announcement',
  READY_NOTICE: 'Ready for Next Level',
  CUSTOM: 'Custom'
};

// Sheet names for communication storage
const COMMUNICATION_SHEETS = {
  HUB: 'CommunicationsHub',
  LOG: 'CommunicationLog',
  TEMPLATES: 'EmailTemplates'
};

// Email template placeholders
const PLACEHOLDERS = {
  STUDENT_NAME: '{{student_name}}',
  PARENT_NAME: '{{parent_name}}',
  CLASS_NAME: '{{class_name}}',
  INSTRUCTOR_NAME: '{{instructor_name}}',
  DAY: '{{day}}',
  TIME: '{{time}}',
  LOCATION: '{{location}}',
  START_DATE: '{{start_date}}',
  END_DATE: '{{end_date}}',
  SESSION_NAME: '{{session_name}}',
  LEVEL: '{{level}}',
  NEXT_LEVEL: '{{next_level}}'
};

/**
 * Creates a communications hub sheet for managing communications
 * 
 * @returns Success status
 */
function createCommunicationsHub() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Creating communications hub', 'INFO', 'createCommunicationsHub');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hubSheet = ss.getSheetByName(COMMUNICATION_SHEETS.HUB);
    
    // Create the sheet if it doesn't exist
    if (!hubSheet) {
      hubSheet = ss.insertSheet(COMMUNICATION_SHEETS.HUB);
      
      // Set up basic structure
      hubSheet.getRange('A1:E1').merge()
        .setValue('YSL v6 Communications Hub')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setBackground('#4285F4')
        .setFontColor('white');
        
      // Instructions
      hubSheet.getRange('A2:E5').merge()
        .setValue('This Communications Hub allows you to create and send communications to parents and students. ' +
                 'Select a communication type, enter the recipient information, and compose your message. ' +
                 'Click "Send" to send the communication.')
        .setWrap(true);
      
      // Communication selection section
      hubSheet.getRange('A7').setValue('Communication Type:').setFontWeight('bold');
      hubSheet.getRange('B7').setValue('Class Email');
      
      // Set up data validation for communication type
      const types = Object.values(COMMUNICATION_TYPES);
      const typeRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(types, true)
        .build();
      hubSheet.getRange('B7').setDataValidation(typeRule);
      
      // Class selection section
      hubSheet.getRange('A9').setValue('Select Class:').setFontWeight('bold');
      createClassDropdown(hubSheet, 'B9');
      
      // Template selection section
      hubSheet.getRange('A11').setValue('Email Template:').setFontWeight('bold');
      hubSheet.getRange('B11').setValue('Welcome Email');
      
      // Set up templates dropdown (later populated by updateTemplateDropdown)
      
      // Subject line
      hubSheet.getRange('A13').setValue('Subject:').setFontWeight('bold');
      hubSheet.getRange('B13:E13').merge()
        .setValue('YSL Swim Lessons Information');
      
      // Message content
      hubSheet.getRange('A15').setValue('Message:').setFontWeight('bold');
      
      hubSheet.getRange('A16:E30').merge()
        .setValue('Dear {{parent_name}},\n\n' +
                 'Welcome to YMCA Swim Lessons! {{student_name}} is registered for {{class_name}} ' +
                 'with {{instructor_name}} on {{day}} at {{time}}.\n\n' +
                 'Classes begin on {{start_date}} and end on {{end_date}}.\n\n' +
                 'Please let us know if you have any questions!\n\n' +
                 'Sincerely,\nYMCA Aquatics Team')
        .setWrap(true);
      
      // Send button visual representation
      hubSheet.getRange('C33:E33').merge()
        .setValue('SEND COMMUNICATION')
        .setBackground('#4285F4')
        .setFontColor('white')
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      // Note about sending
      hubSheet.getRange('A35:E35').merge()
        .setValue('Note: To send the communication, select "Send Selected Communication" from the YSL Hub menu.')
        .setFontStyle('italic')
        .setWrap(true);
      
      // Format sheet
      hubSheet.setColumnWidth(1, 150);
      hubSheet.setColumnWidth(2, 200);
      hubSheet.setColumnWidth(3, 150);
      hubSheet.setColumnWidth(4, 150);
      hubSheet.setColumnWidth(5, 150);
      
      // Create template sheet and log sheet if they don't exist
      createEmailTemplatesSheet(ss);
      createCommunicationLog();
      
      // Update template dropdown
      updateTemplateDropdown(hubSheet);
    }
    
    // Ensure the hub sheet is visible and active
    hubSheet.showSheet();
    hubSheet.activate();
    
    SpreadsheetApp.getUi().alert(
      'Communications Hub Created',
      'The Communications Hub has been created. You can now use it to send communications to parents and students.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createCommunicationsHub', 
        'Error creating communications hub. Please try again or contact support.');
    } else {
      Logger.log(`Error creating communications hub: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create communications hub: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Creates a communication log sheet for tracking communications
 * 
 * @returns Success status
 */
function createCommunicationLog() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Creating communication log', 'INFO', 'createCommunicationLog');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName(COMMUNICATION_SHEETS.LOG);
    
    // Create the sheet if it doesn't exist
    if (!logSheet) {
      logSheet = ss.insertSheet(COMMUNICATION_SHEETS.LOG);
      
      // Set up headers
      const headers = [
        'Date', 'Time', 'Type', 'Recipients', 'Subject', 'Sender', 'Status', 'Notes'
      ];
      
      logSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      logSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      // Set column widths for better visibility
      logSheet.setColumnWidth(1, 100);  // Date
      logSheet.setColumnWidth(2, 80);   // Time
      logSheet.setColumnWidth(3, 120);  // Type
      logSheet.setColumnWidth(4, 200);  // Recipients
      logSheet.setColumnWidth(5, 250);  // Subject
      logSheet.setColumnWidth(6, 150);  // Sender
      logSheet.setColumnWidth(7, 100);  // Status
      logSheet.setColumnWidth(8, 300);  // Notes
      
      // Freeze header row
      logSheet.setFrozenRows(1);
      
      // Create data validation for status
      const statusRange = logSheet.getRange(2, 7, 100, 1);
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Sent', 'Failed', 'Pending', 'Cancelled'], true)
        .build();
      statusRange.setDataValidation(statusRule);
    }
    
    // Ensure the log sheet is visible and active
    logSheet.showSheet();
    logSheet.activate();
    
    SpreadsheetApp.getUi().alert(
      'Communication Log Created',
      'The Communication Log has been created. This sheet will track all communications sent through the system.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createCommunicationLog', 
        'Error creating communication log. Please try again or contact support.');
    } else {
      Logger.log(`Error creating communication log: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create communication log: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Creates an email templates sheet with predefined templates
 * 
 * @param ss - The active spreadsheet
 * @returns Success status
 */
function createEmailTemplatesSheet(ss) {
  try {
    let templatesSheet = ss.getSheetByName(COMMUNICATION_SHEETS.TEMPLATES);
    
    // Create the sheet if it doesn't exist
    if (!templatesSheet) {
      templatesSheet = ss.insertSheet(COMMUNICATION_SHEETS.TEMPLATES);
      
      // Set up headers
      const headers = [
        'Template Name', 'Type', 'Subject', 'Content'
      ];
      
      templatesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      templatesSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      // Add default templates
      const templates = [
        [
          'Welcome Email', 
          COMMUNICATION_TYPES.WELCOME_EMAIL, 
          'Welcome to YMCA Swim Lessons - {{session_name}}',
          'Dear {{parent_name}},\n\n' +
          'Welcome to YMCA Swim Lessons! {{student_name}} is registered for {{class_name}} ' +
          'with {{instructor_name}} on {{day}} at {{time}}.\n\n' +
          'Classes begin on {{start_date}} and end on {{end_date}}.\n\n' +
          'Please let us know if you have any questions!\n\n' +
          'Sincerely,\nYMCA Aquatics Team'
        ],
        [
          'Class Announcement', 
          COMMUNICATION_TYPES.ANNOUNCEMENT, 
          'Important Announcement - {{class_name}}',
          'Dear Parents,\n\n' +
          'This is an important announcement regarding {{class_name}} on {{day}} at {{time}}.\n\n' +
          '[Add your announcement details here]\n\n' +
          'Thank you for your attention to this matter.\n\n' +
          'Sincerely,\nYMCA Aquatics Team'
        ],
        [
          'Ready for Next Level', 
          COMMUNICATION_TYPES.READY_NOTICE, 
          'Your child is ready for the next level!',
          'Dear {{parent_name}},\n\n' +
          'Congratulations! {{student_name}} has completed all requirements for {{level}} ' +
          'and is ready to move up to {{next_level}}.\n\n' +
          'We appreciate your support in your child\'s swimming journey and look forward ' +
          'to seeing their continued progress in our program.\n\n' +
          'Sincerely,\nYMCA Aquatics Team'
        ],
        [
          'Class Update', 
          COMMUNICATION_TYPES.CLASS_EMAIL, 
          'Class Update - {{class_name}}',
          'Dear Parents,\n\n' +
          'I wanted to provide an update on our {{class_name}} swim lessons:\n\n' +
          '1. Progress: [Add progress details here]\n' +
          '2. Focus Areas: [Add focus details here]\n' +
          '3. Upcoming Activities: [Add upcoming activities here]\n\n' +
          'Please let me know if you have any questions!\n\n' +
          'Sincerely,\n{{instructor_name}}'
        ]
      ];
      
      // Add templates to sheet
      templatesSheet.getRange(2, 1, templates.length, templates[0].length).setValues(templates);
      
      // Format sheet
      templatesSheet.setColumnWidth(1, 150);  // Template Name
      templatesSheet.setColumnWidth(2, 120);  // Type
      templatesSheet.setColumnWidth(3, 250);  // Subject
      templatesSheet.setColumnWidth(4, 450);  // Content
      
      // Freeze header row
      templatesSheet.setFrozenRows(1);
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error creating email templates sheet: ${error.message}`);
    return false;
  }
}

/**
 * Creates a dropdown for class selection
 * 
 * @param sheet - The sheet to add the dropdown to
 * @param cellRef - The cell reference for the dropdown
 */
function createClassDropdown(sheet, cellRef) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      // No classes sheet, create an empty dropdown
      const emptyRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['No classes available'], true)
        .build();
      sheet.getRange(cellRef).setDataValidation(emptyRule);
      return;
    }
    
    // Get class data
    const classData = classesSheet.getDataRange().getValues();
    
    // Skip header row
    if (classData.length <= 1) {
      const emptyRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['No classes available'], true)
        .build();
      sheet.getRange(cellRef).setDataValidation(emptyRule);
      return;
    }
    
    // Extract class information for the dropdown
    const classOptions = [];
    for (let i = 1; i < classData.length; i++) {
      // Assuming columns: Class ID, Class Name, Day, Time, Instructor
      const className = classData[i][1];
      const day = classData[i][4];
      const time = classData[i][5];
      const instructor = classData[i][3];
      
      // Add to options if the class has a name
      if (className) {
        classOptions.push(`${className} (${day} ${time}, ${instructor})`);
      }
    }
    
    // Create the dropdown rule
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(classOptions.length > 0 ? classOptions : ['No classes available'], true)
      .build();
    
    // Apply the rule to the cell
    sheet.getRange(cellRef).setDataValidation(rule);
  } catch (error) {
    Logger.log(`Error creating class dropdown: ${error.message}`);
  }
}

/**
 * Updates the template dropdown in the communications hub
 * 
 * @param hubSheet - The communications hub sheet
 */
function updateTemplateDropdown(hubSheet) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templatesSheet = ss.getSheetByName(COMMUNICATION_SHEETS.TEMPLATES);
    
    if (!templatesSheet) {
      return;
    }
    
    // Get template data
    const templateData = templatesSheet.getDataRange().getValues();
    
    // Skip header row
    if (templateData.length <= 1) {
      return;
    }
    
    // Extract template names for the dropdown
    const templateOptions = [];
    for (let i = 1; i < templateData.length; i++) {
      const templateName = templateData[i][0];
      
      if (templateName) {
        templateOptions.push(templateName);
      }
    }
    
    // Add "Custom" option
    templateOptions.push('Custom');
    
    // Create the dropdown rule
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(templateOptions, true)
      .build();
    
    // Apply the rule to the template selection cell
    hubSheet.getRange('B11').setDataValidation(rule);
  } catch (error) {
    Logger.log(`Error updating template dropdown: ${error.message}`);
  }
}

/**
 * Sends the selected communication from the communications hub
 * 
 * @returns Success status
 */
function sendSelectedCommunication() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Sending selected communication', 'INFO', 'sendSelectedCommunication');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hubSheet = ss.getSheetByName(COMMUNICATION_SHEETS.HUB);
    
    if (!hubSheet) {
      SpreadsheetApp.getUi().alert(
        'Communications Hub Not Found',
        'The Communications Hub sheet is missing. Please create it first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get communication details
    const communicationType = hubSheet.getRange('B7').getValue();
    const selectedClass = hubSheet.getRange('B9').getValue();
    const subject = hubSheet.getRange('B13').getValue();
    const message = hubSheet.getRange('A16:E30').getMergedRanges()[0].getValue();
    
    if (!subject || !message) {
      SpreadsheetApp.getUi().alert(
        'Missing Information',
        'Please provide both a subject and message for the communication.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Check class selection for relevant communication types
    if ((communicationType === COMMUNICATION_TYPES.CLASS_EMAIL || 
         communicationType === COMMUNICATION_TYPES.ANNOUNCEMENT || 
         communicationType === COMMUNICATION_TYPES.WELCOME_EMAIL) && 
        (!selectedClass || selectedClass === 'No classes available')) {
      SpreadsheetApp.getUi().alert(
        'Class Selection Required',
        'Please select a class for this type of communication.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get recipients based on communication type and class
    const recipients = getRecipients(communicationType, selectedClass);
    
    if (recipients.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Recipients',
        'No recipients were found for this communication. Please check class enrollment or modify your selection.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Get config values for replacements
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || '';
    
    // Confirm sending
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Send Communication',
      `You are about to send this communication to ${recipients.length} recipients. Continue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return false;
    }
    
    // Get class details for replacements (if applicable)
    let classDetails = null;
    if (selectedClass && selectedClass !== 'No classes available') {
      classDetails = getClassDetails(selectedClass);
    }
    
    // Process and send emails
    let successCount = 0;
    let failCount = 0;
    const errors = [];
    
    for (const recipient of recipients) {
      try {
        // Personalize message and subject for each recipient
        const personalizedSubject = personalizeText(subject, recipient, classDetails, sessionName);
        const personalizedMessage = personalizeText(message, recipient, classDetails, sessionName);
        
        // Send email
        GmailApp.sendEmail(
          recipient.email,
          personalizedSubject,
          personalizedMessage,
          {
            name: 'YMCA Swim Lessons'
          }
        );
        
        successCount++;
      } catch (sendError) {
        failCount++;
        errors.push(`Error sending to ${recipient.email}: ${sendError.message}`);
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error sending to ${recipient.email}: ${sendError.message}`, 'ERROR', 'sendSelectedCommunication');
        }
      }
    }
    
    // Log the communication
    logCommunication(communicationType, recipients.length, subject, successCount, failCount, errors.join('; '));
    
    // Show results
    if (failCount > 0) {
      ui.alert(
        'Sending Complete with Errors',
        `Successfully sent to ${successCount} of ${recipients.length} recipients. Failed: ${failCount}.\n\nSee Communication Log for details.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Sending Complete',
        `Successfully sent to all ${recipients.length} recipients.`,
        ui.ButtonSet.OK
      );
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendSelectedCommunication', 
        'Error sending communication. Please try again or contact support.');
    } else {
      Logger.log(`Error sending communication: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to send communication: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Gets recipients based on communication type and class
 * 
 * @param communicationType - The type of communication
 * @param selectedClass - The selected class
 * @returns Array of recipient objects with email and replacement data
 */
function getRecipients(communicationType, selectedClass) {
  try {
    const recipients = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    if (communicationType === COMMUNICATION_TYPES.CLASS_EMAIL || 
        communicationType === COMMUNICATION_TYPES.ANNOUNCEMENT || 
        communicationType === COMMUNICATION_TYPES.WELCOME_EMAIL) {
      
      // Extract class ID from the selected class text
      const classId = extractClassId(selectedClass);
      
      if (!classId) {
        return [];
      }
      
      // Get roster data
      const rosterSheet = ss.getSheetByName('Roster');
      if (!rosterSheet) {
        return [];
      }
      
      const rosterData = rosterSheet.getDataRange().getValues();
      
      // Skip header row
      for (let i = 1; i < rosterData.length; i++) {
        const row = rosterData[i];
        const studentClassId = row[1];
        
        // Only include students in the selected class with an email
        if (studentClassId === classId && row[7]) { // Assuming column 7 is parent email
          recipients.push({
            email: row[7],
            student_name: `${row[2]} ${row[3]}`, // First + Last name
            parent_name: 'Parent', // We don't have parent names in the roster
            level: row[5] || ''
          });
        }
      }
    } else if (communicationType === COMMUNICATION_TYPES.READY_NOTICE) {
      // For ready notices, we need to filter students based on assessment data
      // This is simplified; real implementation would be more complex
      
      const readyStudents = getStudentsReadyForNextLevel();
      
      for (const student of readyStudents) {
        if (student.parentEmail) {
          recipients.push({
            email: student.parentEmail,
            student_name: student.studentName,
            parent_name: 'Parent',
            level: student.level,
            next_level: student.nextLevel
          });
        }
      }
    } else if (communicationType === COMMUNICATION_TYPES.CUSTOM) {
      // For custom communications, we might use a different approach
      // Example: Manual recipient entry in a separate range
      const customRecipientsRange = ss.getSheetByName(COMMUNICATION_SHEETS.HUB).getRange('A35:B40');
      const customRecipientsData = customRecipientsRange.getValues();
      
      for (const row of customRecipientsData) {
        if (row[0] && row[1] && row[1].includes('@')) {
          recipients.push({
            email: row[1],
            student_name: row[0],
            parent_name: 'Parent'
          });
        }
      }
    }
    
    return recipients;
  } catch (error) {
    Logger.log(`Error getting recipients: ${error.message}`);
    return [];
  }
}

/**
 * Extracts class ID from the selected class text
 * 
 * @param selectedClass - The selected class text
 * @returns The class ID
 */
function extractClassId(selectedClass) {
  try {
    // Get classes data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      return null;
    }
    
    const classData = classesSheet.getDataRange().getValues();
    
    // Skip header row
    for (let i = 1; i < classData.length; i++) {
      const className = classData[i][1];
      const day = classData[i][4];
      const time = classData[i][5];
      const instructor = classData[i][3];
      
      // Check if this is the selected class
      const classText = `${className} (${day} ${time}, ${instructor})`;
      
      if (classText === selectedClass) {
        return classData[i][0]; // Return the class ID
      }
    }
    
    return null;
  } catch (error) {
    Logger.log(`Error extracting class ID: ${error.message}`);
    return null;
  }
}

/**
 * Gets class details for the selected class
 * 
 * @param selectedClass - The selected class text
 * @returns Object with class details
 */
function getClassDetails(selectedClass) {
  try {
    // Get classes data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      return null;
    }
    
    const classData = classesSheet.getDataRange().getValues();
    
    // Skip header row
    for (let i = 1; i < classData.length; i++) {
      const className = classData[i][1];
      const level = classData[i][2];
      const instructor = classData[i][3];
      const day = classData[i][4];
      const time = classData[i][5];
      const startDate = classData[i][6];
      const endDate = classData[i][7];
      const location = classData[i][8];
      
      // Check if this is the selected class
      const classText = `${className} (${day} ${time}, ${instructor})`;
      
      if (classText === selectedClass) {
        return {
          class_name: className,
          level: level,
          instructor_name: instructor,
          day: day,
          time: time,
          start_date: formatDate(startDate),
          end_date: formatDate(endDate),
          location: location
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log(`Error getting class details: ${error.message}`);
    return null;
  }
}

/**
 * Formats a date for display
 * 
 * @param date - The date to format
 * @returns Formatted date string
 */
function formatDate(date) {
  if (!date) return '';
  
  try {
    if (typeof date === 'string') {
      date = new Date(date);
    }
    
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MMMM d, yyyy');
  } catch (error) {
    return date.toString();
  }
}

/**
 * Personalizes text with recipient and class information
 * 
 * @param text - The text to personalize
 * @param recipient - The recipient object
 * @param classDetails - The class details object
 * @param sessionName - The session name
 * @returns Personalized text
 */
function personalizeText(text, recipient, classDetails, sessionName) {
  let result = text;
  
  // Replace recipient placeholders
  if (recipient) {
    result = result.replace(new RegExp(PLACEHOLDERS.STUDENT_NAME, 'g'), recipient.student_name || '');
    result = result.replace(new RegExp(PLACEHOLDERS.PARENT_NAME, 'g'), recipient.parent_name || '');
    result = result.replace(new RegExp(PLACEHOLDERS.LEVEL, 'g'), recipient.level || '');
    result = result.replace(new RegExp(PLACEHOLDERS.NEXT_LEVEL, 'g'), recipient.next_level || '');
  }
  
  // Replace class placeholders
  if (classDetails) {
    result = result.replace(new RegExp(PLACEHOLDERS.CLASS_NAME, 'g'), classDetails.class_name || '');
    result = result.replace(new RegExp(PLACEHOLDERS.INSTRUCTOR_NAME, 'g'), classDetails.instructor_name || '');
    result = result.replace(new RegExp(PLACEHOLDERS.DAY, 'g'), classDetails.day || '');
    result = result.replace(new RegExp(PLACEHOLDERS.TIME, 'g'), classDetails.time || '');
    result = result.replace(new RegExp(PLACEHOLDERS.LOCATION, 'g'), classDetails.location || '');
    result = result.replace(new RegExp(PLACEHOLDERS.START_DATE, 'g'), classDetails.start_date || '');
    result = result.replace(new RegExp(PLACEHOLDERS.END_DATE, 'g'), classDetails.end_date || '');
  }
  
  // Replace session name
  result = result.replace(new RegExp(PLACEHOLDERS.SESSION_NAME, 'g'), sessionName || '');
  
  return result;
}

/**
 * Logs a communication in the communication log
 * 
 * @param type - The type of communication
 * @param recipientCount - The number of recipients
 * @param subject - The subject line
 * @param successCount - Number of successful sends
 * @param failCount - Number of failed sends
 * @param errorDetails - Details of any errors
 */
function logCommunication(type, recipientCount, subject, successCount, failCount, errorDetails) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheet = ss.getSheetByName(COMMUNICATION_SHEETS.LOG);
    
    if (!logSheet) {
      return;
    }
    
    const now = new Date();
    const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const time = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
    
    const status = failCount > 0 ? 'Partial' : 'Sent';
    const notes = failCount > 0 ? `${successCount} sent, ${failCount} failed. Errors: ${errorDetails}` : `${successCount} sent successfully`;
    
    // Prepare log entry
    const logEntry = [
      date,
      time,
      type,
      `${recipientCount} recipients`,
      subject,
      Session.getEffectiveUser().getEmail(),
      status,
      notes
    ];
    
    // Add to log at the top (after header row)
    logSheet.insertRowAfter(1);
    logSheet.getRange(2, 1, 1, logEntry.length).setValues([logEntry]);
  } catch (error) {
    Logger.log(`Error logging communication: ${error.message}`);
  }
}

/**
 * Gets students who are ready for the next level
 * 
 * @returns Array of student objects with assessment data
 */
function getStudentsReadyForNextLevel() {
  // This is a simplified implementation
  // Real implementation would analyze assessment data
  
  const readyStudents = [];
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    
    if (!trackerSheet) {
      return readyStudents;
    }
    
    // Get class information
    const classLevel = trackerSheet.getRange('B4').getValue();
    const nextLevel = getNextLevel(classLevel);
    
    if (!nextLevel) {
      return readyStudents;
    }
    
    // Get student data from the tracker
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
      return readyStudents;
    }
    
    // Get skills from header row
    const skills = [];
    for (let i = 2; i < dataRange[studentStartRow].length; i++) {
      if (dataRange[studentStartRow][i]) {
        skills.push(dataRange[studentStartRow][i]);
      }
    }
    
    // Get roster data for email addresses
    const rosterSheet = ss.getSheetByName('Roster');
    let rosterData = [];
    if (rosterSheet) {
      rosterData = rosterSheet.getDataRange().getValues();
    }
    
    // Analyze each student
    for (let i = studentStartRow + 1; i < dataRange.length; i++) {
      const row = dataRange[i];
      const studentId = row[0];
      const studentName = row[1];
      
      if (!studentId || !studentName) continue;
      
      // Count completed skills
      let completedSkills = 0;
      let totalSkills = 0;
      
      for (let j = 2; j < row.length && j - 2 < skills.length; j++) {
        if (skills[j - 2]) {
          totalSkills++;
          if (row[j] === 'P' || row[j] === 'Proficient' || row[j] === 'Pass') {
            completedSkills++;
          }
        }
      }
      
      // Check if student is ready (completed at least 80% of skills)
      if (totalSkills > 0 && completedSkills / totalSkills >= 0.8) {
        // Find parent email from roster
        let parentEmail = '';
        for (let k = 1; k < rosterData.length; k++) {
          if (rosterData[k][0] === studentId) {
            parentEmail = rosterData[k][7]; // Assuming column 7 is parent email
            break;
          }
        }
        
        readyStudents.push({
          studentId: studentId,
          studentName: studentName,
          level: classLevel,
          nextLevel: nextLevel,
          completedSkills: completedSkills,
          totalSkills: totalSkills,
          parentEmail: parentEmail
        });
      }
    }
    
    return readyStudents;
  } catch (error) {
    Logger.log(`Error finding ready students: ${error.message}`);
    return [];
  }
}

/**
 * Gets the next level based on current level
 * 
 * @param currentLevel - The current level
 * @returns The next level
 */
function getNextLevel(currentLevel) {
  if (!currentLevel) return '';
  
  // Common YMCA swim lesson level progression
  const levelProgression = {
    'Water Discovery': 'Water Exploration',
    'Water Exploration': 'Water Acclimation',
    'Water Acclimation': 'Water Movement',
    'Water Movement': 'Water Stamina',
    'Water Stamina': 'Stroke Introduction',
    'Stroke Introduction': 'Stroke Development',
    'Stroke Development': 'Stroke Mechanics',
    'Stroke Mechanics': 'Pathways'
  };
  
  return levelProgression[currentLevel] || '';
}

/**
 * Sends welcome emails to parents of students in a class
 * 
 * @returns Success status
 */
function sendWelcomeEmails() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Sending welcome emails', 'INFO', 'sendWelcomeEmails');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Get class selection from user
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      ui.alert(
        'No Classes Found',
        'No classes were found in the system. Please add classes first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    const classData = classesSheet.getDataRange().getValues();
    
    if (classData.length <= 1) {
      ui.alert(
        'No Classes Found',
        'No classes were found in the system. Please add classes first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Extract class options for display
    const classOptions = [];
    for (let i = 1; i < classData.length; i++) {
      const className = classData[i][1];
      const day = classData[i][4];
      const time = classData[i][5];
      const instructor = classData[i][3];
      
      if (className) {
        classOptions.push(`${className} (${day} ${time}, ${instructor})`);
      }
    }
    
    // Create a temporary UI for selection
    let tempSheet = ss.getSheetByName('TempClassSelection');
    if (!tempSheet) {
      tempSheet = ss.insertSheet('TempClassSelection');
    } else {
      tempSheet.clear();
    }
    
    // Add instructions and classes list
    tempSheet.getRange('A1:D1').merge()
      .setValue('Select Classes for Welcome Emails')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
      
    tempSheet.getRange('A2:D4').merge()
      .setValue('Check the boxes next to the classes you want to send welcome emails to. ' +
               'Then, click "Send Welcome Emails" from the YSL Hub menu again.')
      .setWrap(true);
    
    // Add class list with checkboxes
    tempSheet.getRange('A6').setValue('Select');
    tempSheet.getRange('B6').setValue('Class');
    tempSheet.getRange('C6').setValue('Students');
    tempSheet.getRange('D6').setValue('Email Status');
    
    tempSheet.getRange('A6:D6')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add classes with checkboxes
    for (let i = 0; i < classOptions.length; i++) {
      tempSheet.getRange(7 + i, 1).insertCheckboxes();
      tempSheet.getRange(7 + i, 2).setValue(classOptions[i]);
      
      // Count students in this class
      const classId = classData[i + 1][0];
      const studentCount = countStudentsInClass(classId);
      tempSheet.getRange(7 + i, 3).setValue(studentCount);
      tempSheet.getRange(7 + i, 4).setValue('Not Sent');
    }
    
    // Format sheet
    tempSheet.setColumnWidth(1, 80);
    tempSheet.setColumnWidth(2, 300);
    tempSheet.setColumnWidth(3, 100);
    tempSheet.setColumnWidth(4, 100);
    
    // Activate sheet
    tempSheet.activate();
    
    // Get the template
    const welcomeEmailTemplate = getEmailTemplate('Welcome Email');
    
    // If we don't have any checked boxes, just show the sheet and return
    const allCheckboxes = tempSheet.getRange(7, 1, classOptions.length, 1).getValues();
    let anyChecked = false;
    for (const row of allCheckboxes) {
      if (row[0] === true) {
        anyChecked = true;
        break;
      }
    }
    
    if (!anyChecked) {
      ui.alert(
        'Select Classes',
        'Please select the classes you want to send welcome emails to, then click "Send Welcome Emails" again.',
        ui.ButtonSet.OK
      );
      return true;
    }
    
    // Send emails for checked classes
    let totalSent = 0;
    let totalFailed = 0;
    const errors = [];
    
    // Get session name for replacements
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || '';
    
    // Process each class
    for (let i = 0; i < classOptions.length; i++) {
      // Check if this class is selected
      const isSelected = tempSheet.getRange(7 + i, 1).getValue();
      
      if (isSelected) {
        const classId = classData[i + 1][0];
        const classDetails = {
          class_name: classData[i + 1][1],
          level: classData[i + 1][2],
          instructor_name: classData[i + 1][3],
          day: classData[i + 1][4],
          time: classData[i + 1][5],
          start_date: formatDate(classData[i + 1][6]),
          end_date: formatDate(classData[i + 1][7]),
          location: classData[i + 1][8]
        };
        
        // Get recipients for this class
        const recipients = getClassRecipients(classId);
        
        // Update status cell
        tempSheet.getRange(7 + i, 4).setValue(`Sending to ${recipients.length}...`);
        
        // Send emails
        let classSent = 0;
        let classFailed = 0;
        
        for (const recipient of recipients) {
          try {
            // Personalize message and subject
            const personalizedSubject = personalizeText(welcomeEmailTemplate.subject, recipient, classDetails, sessionName);
            const personalizedMessage = personalizeText(welcomeEmailTemplate.content, recipient, classDetails, sessionName);
            
            // Send email
            GmailApp.sendEmail(
              recipient.email,
              personalizedSubject,
              personalizedMessage,
              {
                name: 'YMCA Swim Lessons'
              }
            );
            
            classSent++;
            totalSent++;
          } catch (sendError) {
            classFailed++;
            totalFailed++;
            errors.push(`Error sending to ${recipient.email}: ${sendError.message}`);
            
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`Error sending welcome email to ${recipient.email}: ${sendError.message}`, 'ERROR', 'sendWelcomeEmails');
            }
          }
        }
        
        // Update status cell
        if (classFailed > 0) {
          tempSheet.getRange(7 + i, 4).setValue(`Sent ${classSent}, Failed ${classFailed}`).setBackground('#FFC107');
        } else if (classSent > 0) {
          tempSheet.getRange(7 + i, 4).setValue(`Sent ${classSent}`).setBackground('#4CAF50');
        } else {
          tempSheet.getRange(7 + i, 4).setValue('No recipients').setBackground('#F44336');
        }
        
        // Log the communication
        logCommunication(
          COMMUNICATION_TYPES.WELCOME_EMAIL,
          recipients.length,
          welcomeEmailTemplate.subject,
          classSent,
          classFailed,
          classFailed > 0 ? `Errors sending to class ${classDetails.class_name}` : ''
        );
      }
    }
    
    // Show results
    if (totalFailed > 0) {
      ui.alert(
        'Sending Complete with Errors',
        `Successfully sent ${totalSent} welcome emails. Failed: ${totalFailed}.\n\nSee Communication Log for details.`,
        ui.ButtonSet.OK
      );
    } else if (totalSent > 0) {
      ui.alert(
        'Sending Complete',
        `Successfully sent ${totalSent} welcome emails.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'No Emails Sent',
        'No welcome emails were sent. Please check that the selected classes have students with valid email addresses.',
        ui.ButtonSet.OK
      );
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendWelcomeEmails', 
        'Error sending welcome emails. Please try again or contact support.');
    } else {
      Logger.log(`Error sending welcome emails: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to send welcome emails: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Counts the number of students in a class
 * 
 * @param classId - The class ID
 * @returns Number of students
 */
function countStudentsInClass(classId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName('Roster');
    
    if (!rosterSheet) {
      return 0;
    }
    
    const rosterData = rosterSheet.getDataRange().getValues();
    let count = 0;
    
    // Skip header row
    for (let i = 1; i < rosterData.length; i++) {
      if (rosterData[i][1] === classId) { // Assuming column 1 is class ID
        count++;
      }
    }
    
    return count;
  } catch (error) {
    Logger.log(`Error counting students: ${error.message}`);
    return 0;
  }
}

/**
 * Gets recipients for a specific class
 * 
 * @param classId - The class ID
 * @returns Array of recipient objects
 */
function getClassRecipients(classId) {
  try {
    const recipients = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName('Roster');
    
    if (!rosterSheet) {
      return recipients;
    }
    
    const rosterData = rosterSheet.getDataRange().getValues();
    
    // Skip header row
    for (let i = 1; i < rosterData.length; i++) {
      const row = rosterData[i];
      const studentClassId = row[1];
      
      // Only include students in the selected class with an email
      if (studentClassId === classId && row[7]) { // Assuming column 7 is parent email
        recipients.push({
          email: row[7],
          student_name: `${row[2]} ${row[3]}`, // First + Last name
          parent_name: 'Parent', // We don't have parent names in the roster
          level: row[5] || ''
        });
      }
    }
    
    return recipients;
  } catch (error) {
    Logger.log(`Error getting class recipients: ${error.message}`);
    return [];
  }
}

/**
 * Gets an email template by name
 * 
 * @param templateName - The name of the template
 * @returns Template object with subject and content
 */
function getEmailTemplate(templateName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templatesSheet = ss.getSheetByName(COMMUNICATION_SHEETS.TEMPLATES);
    
    if (!templatesSheet) {
      return {
        subject: 'YMCA Swim Lessons',
        content: 'Welcome to YMCA Swim Lessons!'
      };
    }
    
    const templateData = templatesSheet.getDataRange().getValues();
    
    // Skip header row, find matching template
    for (let i = 1; i < templateData.length; i++) {
      if (templateData[i][0] === templateName) {
        return {
          subject: templateData[i][2],
          content: templateData[i][3]
        };
      }
    }
    
    // Return default if not found
    return {
      subject: 'YMCA Swim Lessons',
      content: 'Welcome to YMCA Swim Lessons!'
    };
  } catch (error) {
    Logger.log(`Error getting email template: ${error.message}`);
    return {
      subject: 'YMCA Swim Lessons',
      content: 'Welcome to YMCA Swim Lessons!'
    };
  }
}

/**
 * Tests sending a welcome email to a specified address
 * 
 * @returns Success status
 */
function testWelcomeEmail() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Testing welcome email', 'INFO', 'testWelcomeEmail');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Get test email address
    const emailResult = ui.prompt(
      'Test Welcome Email',
      'Enter the email address to send the test to:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (emailResult.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    const testEmail = emailResult.getResponseText().trim();
    
    if (!testEmail || !testEmail.includes('@')) {
      ui.alert(
        'Invalid Email',
        'Please enter a valid email address.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Get the welcome email template
    const welcomeEmailTemplate = getEmailTemplate('Welcome Email');
    
    // Get example class data for replacements
    const exampleClass = {
      class_name: 'Water Acclimation',
      level: 'Preschool',
      instructor_name: 'Jane Smith',
      day: 'Monday',
      time: '4:00 PM',
      start_date: 'June 5, 2025',
      end_date: 'July 10, 2025',
      location: 'Main Pool'
    };
    
    // Get example student data
    const exampleStudent = {
      student_name: 'Alex Johnson',
      parent_name: 'Parent',
      level: 'Preschool'
    };
    
    // Get session name
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || 'Summer 2025';
    
    // Personalize message and subject
    const subject = personalizeText(welcomeEmailTemplate.subject, exampleStudent, exampleClass, sessionName);
    const message = personalizeText(welcomeEmailTemplate.content, exampleStudent, exampleClass, sessionName);
    
    // Show preview
    const previewResult = ui.alert(
      'Preview Test Email',
      `This test email will be sent to: ${testEmail}\n\nSubject: ${subject}\n\nMessage:\n${message}\n\nContinue?`,
      ui.ButtonSet.YES_NO
    );
    
    if (previewResult !== ui.Button.YES) {
      return false;
    }
    
    // Send the test email
    GmailApp.sendEmail(
      testEmail,
      subject,
      message,
      {
        name: 'YMCA Swim Lessons'
      }
    );
    
    // Log the test
    logCommunication(
      'Test Email',
      1,
      subject,
      1,
      0,
      `Test welcome email sent to ${testEmail}`
    );
    
    ui.alert(
      'Test Email Sent',
      `A test welcome email has been sent to ${testEmail}.`,
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'testWelcomeEmail', 
        'Error sending test welcome email. Please try again or contact support.');
    } else {
      Logger.log(`Error sending test welcome email: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to send test welcome email: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Sends emails to all parents/contacts of students in a class
 * 
 * @returns Success status
 */
function emailClassParticipants() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Emailing class participants', 'INFO', 'emailClassParticipants');
    }
    
    // Create communications hub if it doesn't exist and show it
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hubSheet = ss.getSheetByName(COMMUNICATION_SHEETS.HUB);
    
    if (!hubSheet) {
      createCommunicationsHub();
      return true;
    }
    
    // Set up the hub for class email
    hubSheet.getRange('B7').setValue(COMMUNICATION_TYPES.CLASS_EMAIL);
    
    // Select the first available class
    const classesSheet = ss.getSheetByName('Classes');
    if (classesSheet && classesSheet.getLastRow() > 1) {
      createClassDropdown(hubSheet, 'B9');
    }
    
    // Show the hub
    hubSheet.activate();
    
    SpreadsheetApp.getUi().alert(
      'Email Class Participants',
      'To email class participants, select a class from the dropdown, customize your message, and click "Send Selected Communication" from the YSL Hub menu.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'emailClassParticipants', 
        'Error setting up class email. Please try again or contact support.');
    } else {
      Logger.log(`Error setting up class email: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to set up class email: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Sends class announcements to all classes or selected classes
 * 
 * @returns Success status
 */
function sendClassAnnouncements() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Sending class announcements', 'INFO', 'sendClassAnnouncements');
    }
    
    // Create communications hub if it doesn't exist and show it
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hubSheet = ss.getSheetByName(COMMUNICATION_SHEETS.HUB);
    
    if (!hubSheet) {
      createCommunicationsHub();
      return true;
    }
    
    // Set up the hub for class announcements
    hubSheet.getRange('B7').setValue(COMMUNICATION_TYPES.ANNOUNCEMENT);
    
    // Use the Class Announcement template
    updateTemplateDropdown(hubSheet);
    hubSheet.getRange('B11').setValue('Class Announcement');
    
    // Update subject and content from template
    const template = getEmailTemplate('Class Announcement');
    hubSheet.getRange('B13').setValue(template.subject);
    hubSheet.getRange('A16:E30').getMergedRanges()[0].setValue(template.content);
    
    // Show the hub
    hubSheet.activate();
    
    SpreadsheetApp.getUi().alert(
      'Send Class Announcements',
      'To send class announcements, select a class from the dropdown, customize your message, and click "Send Selected Communication" from the YSL Hub menu.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendClassAnnouncements', 
        'Error setting up class announcements. Please try again or contact support.');
    } else {
      Logger.log(`Error setting up class announcements: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to set up class announcements: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Sends notifications to parents of students who are ready for the next level
 * 
 * @returns Success status
 */
function sendReadyAnnouncements() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Sending ready announcements', 'INFO', 'sendReadyAnnouncements');
    }
    
    // Create communications hub if it doesn't exist and show it
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hubSheet = ss.getSheetByName(COMMUNICATION_SHEETS.HUB);
    
    if (!hubSheet) {
      createCommunicationsHub();
      return true;
    }
    
    // Set up the hub for ready announcements
    hubSheet.getRange('B7').setValue(COMMUNICATION_TYPES.READY_NOTICE);
    
    // Use the Ready for Next Level template
    updateTemplateDropdown(hubSheet);
    hubSheet.getRange('B11').setValue('Ready for Next Level');
    
    // Update subject and content from template
    const template = getEmailTemplate('Ready for Next Level');
    hubSheet.getRange('B13').setValue(template.subject);
    hubSheet.getRange('A16:E30').getMergedRanges()[0].setValue(template.content);
    
    // Show the hub
    hubSheet.activate();
    
    const readyStudents = getStudentsReadyForNextLevel();
    
    if (readyStudents.length === 0) {
      SpreadsheetApp.getUi().alert(
        'No Students Ready',
        'No students were found who are ready for the next level. Please ensure that you have assessment data in the Group Lesson Tracker sheet.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    SpreadsheetApp.getUi().alert(
      'Send Ready Announcements',
      `${readyStudents.length} students were found who are ready for the next level. Customize your message if needed, and click "Send Selected Communication" from the YSL Hub menu.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendReadyAnnouncements', 
        'Error setting up ready announcements. Please try again or contact support.');
    } else {
      Logger.log(`Error setting up ready announcements: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to set up ready announcements: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

// Global variable export
const CommunicationModule = {
  createCommunicationsHub,
  createCommunicationLog,
  sendSelectedCommunication,
  emailClassParticipants,
  sendClassAnnouncements,
  sendReadyAnnouncements,
  sendWelcomeEmails,
  testWelcomeEmail
};