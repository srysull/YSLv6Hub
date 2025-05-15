/**
 * YSL Hub v2 Communication Module
 * 
 * This module handles communications with class participants and stakeholders,
 * including email notifications, class announcements, and schedule updates.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-14
 */

// Default sender details
const DEFAULT_SENDER = {
  EMAIL: 'ssullivan@penbayymca.org',
  NAME: 'PenBayY - Aquatics'
};

/**
 * Sends an email to all participants in the selected classes.
 * 
 * @return {boolean} Success status
 */
function emailClassParticipants() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting email to class participants', 'INFO', 'emailClassParticipants');
    }
    
    // Get selected classes
    const selectedClasses = getSelectedClasses();
    if (selectedClasses.length === 0) {
      ui.alert(
        'No Classes Selected',
        'Please select at least one class in the Classes sheet by setting the "Select Class" column to "Select".',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Get email subject
    const subjectResponse = ui.prompt(
      'Email Class Participants',
      'Enter the email subject:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (subjectResponse.getSelectedButton() !== ui.Button.OK) return false;
    const subject = subjectResponse.getResponseText().trim();
    
    if (!subject) {
      ui.alert('Error', 'Email subject is required.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get email body
    const bodyResponse = ui.prompt(
      'Email Class Participants',
      'Enter the email message (you can use HTML formatting):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (bodyResponse.getSelectedButton() !== ui.Button.OK) return false;
    const body = bodyResponse.getResponseText().trim();
    
    if (!body) {
      ui.alert('Error', 'Email body is required.', ui.ButtonSet.OK);
      return false;
    }
    
    // Confirm sending
    const totalStudents = selectedClasses.reduce((sum, classInfo) => sum + classInfo.count, 0);
    
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Confirm Email',
        `This will send an email to approximately ${totalStudents} participants across ${selectedClasses.length} classes. Continue?`
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Confirm Email',
        `This will send an email to approximately ${totalStudents} participants across ${selectedClasses.length} classes. Continue?`,
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Process each class and collect recipients
    let recipients = [];
    let classesProcessed = 0;
    let errorCount = 0;
    
    for (const classInfo of selectedClasses) {
      try {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Processing class: ${classInfo.program} (${classInfo.day}, ${classInfo.time})`, 'INFO', 'emailClassParticipants');
        }
        
        const classRecipients = getClassRecipients(classInfo);
        recipients = recipients.concat(classRecipients);
        classesProcessed++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Found ${classRecipients.length} recipients for class`, 'INFO', 'emailClassParticipants');
        }
      } catch (error) {
        errorCount++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error processing class ${classInfo.program}: ${error.message}`, 'ERROR', 'emailClassParticipants');
        } else {
          Logger.log(`Error processing class ${classInfo.program}: ${error.message}`);
        }
      }
    }
    
    // Remove duplicates
    recipients = [...new Set(recipients)];
    
    // If no recipients found
    if (recipients.length === 0) {
      ui.alert(
        'No Recipients Found',
        'No email addresses were found for the selected classes.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Sending email to ${recipients.length} unique recipients`, 'INFO', 'emailClassParticipants');
    }
    
    // Send emails in batches to avoid quota limits
    const batchSize = 50; // Gmail batch sending limit
    const totalBatches = Math.ceil(recipients.length / batchSize);
    let emailsSent = 0;
    let batchErrors = 0;
    
    for (let i = 0; i < totalBatches; i++) {
      const batch = recipients.slice(i * batchSize, (i + 1) * batchSize);
      
      try {
        // Send as BCC to protect privacy
        GmailApp.sendEmail('', subject, '', {
          htmlBody: body,
          bcc: batch.join(','),
          from: DEFAULT_SENDER.EMAIL,
          name: DEFAULT_SENDER.NAME
        });
        
        emailsSent += batch.length;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Sent batch ${i+1}/${totalBatches} successfully`, 'INFO', 'emailClassParticipants');
        }
        
        // Wait between batches to avoid quota limits
        if (i < totalBatches - 1) {
          Utilities.sleep(1000);
        }
      } catch (error) {
        batchErrors++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error sending batch ${i+1}: ${error.message}`, 'ERROR', 'emailClassParticipants');
        } else {
          Logger.log(`Error sending batch ${i+1}: ${error.message}`);
        }
      }
    }
    
    // Show results
    ui.alert(
      'Email Results',
      `Email sent to ${emailsSent} of ${recipients.length} recipients from ${classesProcessed} classes.\n` +
      (errorCount > 0 || batchErrors > 0 ? `${errorCount + batchErrors} errors occurred. Check logs for details.` : ''),
      ui.ButtonSet.OK
    );
    
    // Log summary
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Email sent to ${emailsSent} recipients (${recipients.length} total) from ${classesProcessed} classes`, 'INFO', 'emailClassParticipants');
    } else {
      Logger.log(`Email sent to ${emailsSent} recipients (${recipients.length} total) from ${classesProcessed} classes.`);
    }
    
    return emailsSent > 0;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'emailClassParticipants', 
        'Failed to send emails. Please try again or contact support.');
    } else {
      Logger.log(`Error in emailClassParticipants: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to send emails: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Sends class-specific announcements to participants.
 * 
 * @return {boolean} Success status
 */
function sendClassAnnouncements() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting class announcements setup', 'INFO', 'sendClassAnnouncements');
    }
    
    // Get selected classes
    const selectedClasses = getSelectedClasses();
    if (selectedClasses.length === 0) {
      ui.alert(
        'No Classes Selected',
        'Please select at least one class in the Classes sheet by setting the "Select Class" column to "Select".',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Confirm action
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Send Class Announcements',
        `This will prepare class-specific announcements for ${selectedClasses.length} classes. Continue?`
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Send Class Announcements',
        `This will prepare class-specific announcements for ${selectedClasses.length} classes. Continue?`,
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Create announcements sheet if it doesn't exist
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let announcementsSheet = ss.getSheetByName('Announcements');
    
    if (!announcementsSheet) {
      announcementsSheet = ss.insertSheet('Announcements');
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Created new Announcements sheet', 'INFO', 'sendClassAnnouncements');
      }
      
      // Set up headers
      const headers = ['Class', 'Program', 'Day', 'Time', 'Instructor', 'Student Count', 'Email Subject', 'Message', 'Status', 'Sent Date'];
      announcementsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      announcementsSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      // Freeze header row
      announcementsSheet.setFrozenRows(1);
      
      // Set column widths
      announcementsSheet.setColumnWidth(1, 150);  // Class
      announcementsSheet.setColumnWidth(2, 150);  // Program
      announcementsSheet.setColumnWidth(3, 100);  // Day
      announcementsSheet.setColumnWidth(4, 100);  // Time
      announcementsSheet.setColumnWidth(5, 150);  // Instructor
      announcementsSheet.setColumnWidth(6, 100);  // Student Count
      announcementsSheet.setColumnWidth(7, 200);  // Email Subject
      announcementsSheet.setColumnWidth(8, 400);  // Message
      announcementsSheet.setColumnWidth(9, 100);  // Status
      announcementsSheet.setColumnWidth(10, 150); // Sent Date
    }
    
    // Add selected classes to announcements sheet
    let rowIndex = announcementsSheet.getLastRow() + 1;
    let classesAdded = 0;
    
    for (const classInfo of selectedClasses) {
      // Create class identifier
      const classId = `${classInfo.program} - ${classInfo.day} ${classInfo.time}`;
      
      // Check if class already exists in announcements sheet
      let classExists = false;
      const existingData = announcementsSheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        if (existingData[i][0] === classId) {
          classExists = true;
          
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Class already exists in announcements: ${classId}`, 'INFO', 'sendClassAnnouncements');
          }
          
          break;
        }
      }
      
      // Skip if already exists
      if (classExists) continue;
      
      // Add class to announcements sheet
      const rowData = [
        classId,
        classInfo.program,
        classInfo.day,
        classInfo.time,
        classInfo.instructor,
        classInfo.count,
        `[YSL] ${classInfo.program} Class Update`,
        '',
        'Draft',
        ''
      ];
      
      announcementsSheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      
      // Add data validation for status
      const statusValidation = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Draft', 'Ready', 'Sent'], true)
        .build();
      
      announcementsSheet.getRange(rowIndex, 9).setDataValidation(statusValidation);
      
      rowIndex++;
      classesAdded++;
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Added class to announcements: ${classId}`, 'INFO', 'sendClassAnnouncements');
      }
    }
    
    // Add button for sending announcements
    let sendAnnouncementsButton = null;
    const drawings = announcementsSheet.getDrawings();
    
    for (let i = 0; i < drawings.length; i++) {
      if (drawings[i].getAltDescription && drawings[i].getAltDescription() === 'SendAnnouncementsButton') {
        sendAnnouncementsButton = drawings[i];
        break;
      }
    }
    
    if (!sendAnnouncementsButton) {
      try {
        // Create a text box that looks like a button
        const buttonCell = announcementsSheet.getRange(announcementsSheet.getLastRow() + 2, 1);
        const textBox = announcementsSheet.insertTextBox('Send Ready Announcements');
        
        textBox.setPosition(buttonCell.getRow(), buttonCell.getColumn(), 0, 0);
        textBox.setWidth(200);
        textBox.setHeight(30);
        textBox.getText().setFontSize(12).setFontWeight('bold');
        textBox.setFill('#4285F4');
        textBox.getText().setForegroundColor('#FFFFFF');
        textBox.setBorder(true, true, true, true, true, true, '#3367D6', null);
        
        if (textBox.setAltDescription) {
          textBox.setAltDescription('SendAnnouncementsButton');
        }
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage('Added Send Ready Announcements button', 'INFO', 'sendClassAnnouncements');
        }
      } catch (error) {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error creating Send Announcements button: ${error.message}`, 'WARNING', 'sendClassAnnouncements');
        } else {
          Logger.log(`Error creating Send Announcements button: ${error.message}`);
        }
      }
    }
    
    // Activate the announcements sheet
    announcementsSheet.activate();
    
    // Show instructions
    ui.alert(
      'Announcements Setup',
      'The Announcements sheet has been prepared. Please complete the Message column for each class, then set the Status to "Ready" when you want to send.\n\n' +
      `${classesAdded} new classes were added to the announcements sheet.\n\n` +
      'Use the "Send Ready Announcements" button at the bottom of the sheet when you\'re ready to send.',
      ui.ButtonSet.OK
    );
    
    // Log completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Class announcements setup completed: ${classesAdded} classes added`, 'INFO', 'sendClassAnnouncements');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendClassAnnouncements', 
        'Failed to prepare class announcements. Please try again or contact support.');
    } else {
      Logger.log(`Error in sendClassAnnouncements: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to prepare class announcements: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Sends announcements that are marked as "Ready" in the Announcements sheet.
 * 
 * @return {boolean} Success status
 */
function sendReadyAnnouncements() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting to send ready announcements', 'INFO', 'sendReadyAnnouncements');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const announcementsSheet = ss.getSheetByName('Announcements');
    
    if (!announcementsSheet) {
      ui.alert('Error', 'Announcements sheet not found.', ui.ButtonSet.OK);
      return false;
    }
    
    const announcementData = announcementsSheet.getDataRange().getValues();
    
    if (announcementData.length <= 1) {
      ui.alert('No Announcements', 'No announcements found in the Announcements sheet.', ui.ButtonSet.OK);
      return false;
    }
    
    // Find ready announcements
    const readyAnnouncements = [];
    for (let i = 1; i < announcementData.length; i++) {
      if (announcementData[i][8] === 'Ready') {
        readyAnnouncements.push({
          rowIndex: i + 1,
          classId: announcementData[i][0],
          program: announcementData[i][1],
          day: announcementData[i][2],
          time: announcementData[i][3],
          instructor: announcementData[i][4],
          subject: announcementData[i][6],
          message: announcementData[i][7]
        });
      }
    }
    
    if (readyAnnouncements.length === 0) {
      ui.alert('No Ready Announcements', 'No announcements marked as "Ready" were found.', ui.ButtonSet.OK);
      return false;
    }
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Found ${readyAnnouncements.length} ready announcements`, 'INFO', 'sendReadyAnnouncements');
    }
    
    // Confirm sending
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Send Announcements',
        `This will send ${readyAnnouncements.length} announcements to their respective class participants. Continue?`
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Send Announcements',
        `This will send ${readyAnnouncements.length} announcements to their respective class participants. Continue?`,
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Process each announcement
    let announcementsSent = 0;
    let errorCount = 0;
    
    for (const announcement of readyAnnouncements) {
      try {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Processing announcement for class: ${announcement.classId}`, 'INFO', 'sendReadyAnnouncements');
        }
        
        // Create class info object needed for getClassRecipients
        const classInfo = {
          program: announcement.program,
          day: announcement.day,
          time: announcement.time,
          instructor: announcement.instructor
        };
        
        // Get recipients for the class
        const recipients = getClassRecipients(classInfo);
        
        if (recipients.length === 0) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`No recipients found for class: ${announcement.classId}`, 'WARNING', 'sendReadyAnnouncements');
          }
          
          throw new Error(`No recipients found for class: ${announcement.classId}`);
        }
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Found ${recipients.length} recipients for announcement`, 'INFO', 'sendReadyAnnouncements');
        }
        
        // Send the announcement
        const subject = announcement.subject;
        const message = announcement.message;
        
        // Send as BCC to protect privacy
        GmailApp.sendEmail('', subject, '', {
          htmlBody: message,
          bcc: recipients.join(','),
          from: DEFAULT_SENDER.EMAIL,
          name: DEFAULT_SENDER.NAME
        });
        
        // Update status to Sent
        announcementsSheet.getRange(announcement.rowIndex, 9).setValue('Sent');
        announcementsSheet.getRange(announcement.rowIndex, 10).setValue(new Date());
        
        announcementsSent++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Successfully sent announcement for class: ${announcement.classId}`, 'INFO', 'sendReadyAnnouncements');
        }
      } catch (error) {
        errorCount++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error sending announcement for ${announcement.classId}: ${error.message}`, 'ERROR', 'sendReadyAnnouncements');
        } else {
          Logger.log(`Error sending announcement for ${announcement.classId}: ${error.message}`);
        }
      }
    }
    
    // Show results
    ui.alert(
      'Announcement Results',
      `${announcementsSent} of ${readyAnnouncements.length} announcements were sent successfully.\n` +
      (errorCount > 0 ? `${errorCount} errors occurred. Check logs for details.` : ''),
      ui.ButtonSet.OK
    );
    
    // Log completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Announcements sent: ${announcementsSent} successes, ${errorCount} failures`, 'INFO', 'sendReadyAnnouncements');
    }
    
    return announcementsSent > 0;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendReadyAnnouncements', 
        'Failed to send announcements. Please try again or contact support.');
    } else {
      Logger.log(`Error in sendReadyAnnouncements: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to send announcements: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Gets all selected classes from the Classes sheet.
 * 
 * @return {Array} Array of selected class objects
 */
function getSelectedClasses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Classes sheet not found', 'ERROR', 'getSelectedClasses');
      }
      return [];
    }
    
    // Get all data from Classes sheet
    const classData = classesSheet.getDataRange().getValues();
    
    if (classData.length <= 1) {
      return [];
    }
    
    // Find selected classes
    const selectedClasses = [];
    for (let i = 1; i < classData.length; i++) {
      if (classData[i][0] === 'Select') {
        selectedClasses.push({
          rowIndex: i,
          program: classData[i][1],
          day: classData[i][2],
          time: classData[i][3],
          location: classData[i][4],
          count: classData[i][5],
          instructor: classData[i][6]
        });
      }
    }
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Found ${selectedClasses.length} selected classes`, 'INFO', 'getSelectedClasses');
    }
    
    return selectedClasses;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting selected classes: ${error.message}`, 'ERROR', 'getSelectedClasses');
    } else {
      Logger.log(`Error getting selected classes: ${error.message}`);
    }
    return [];
  }
}

/**
 * Gets email recipients for a specific class.
 * 
 * @param {Object} classInfo - Information about the class
 * @return {Array} Array of email addresses
 */
function getClassRecipients(classInfo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName('Daxko');
    
    if (!daxkoSheet) {
      throw new Error('Daxko sheet not found');
    }
    
    // Get all data from Daxko sheet
    const daxkoData = daxkoSheet.getDataRange().getValues();
    const headers = daxkoData[0];
    
    // Find the column indices for relevant fields
    let programColIndex, dayColIndex, timeColIndex, siteColIndex;
    
    // Use the findColumnIndex function if available, otherwise use indexOf
    if (typeof GlobalFunctions.findColumnIndex === 'function') {
      programColIndex = GlobalFunctions.findColumnIndex(headers, ['program', 'class', 'stage', 'level']);
      dayColIndex = GlobalFunctions.findColumnIndex(headers, ['day', 'days', 'day(s)', 'weekday']);
      timeColIndex = GlobalFunctions.findColumnIndex(headers, ['time', 'session time', 'class time']);
      siteColIndex = GlobalFunctions.findColumnIndex(headers, ['location', 'site', 'facility']);
    } else {
      // Fallback to basic indexOf
      programColIndex = headers.indexOf('Program');
      dayColIndex = headers.indexOf('Day(s) of Week');
      timeColIndex = headers.indexOf('Session Time');
      siteColIndex = headers.indexOf('Site');
    }
    
    // Email columns in Daxko data - find email columns using flexible matching
    const emailColIndices = [];
    
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (header && typeof header === 'string' && 
          (header.toLowerCase().includes('email') || 
           header.toLowerCase().includes('e-mail'))) {
        emailColIndices.push(i);
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Found email column: ${header} at index ${i}`, 'DEBUG', 'getClassRecipients');
        }
      }
    }
    
    // If no email columns found, look specifically for these expected columns
    if (emailColIndices.length === 0) {
      const specificEmailCols = [
        'E-mail',                // Primary email (S)
        'Primary Member Email',  // Primary member email (AK)
        'Secondary Member Email' // Secondary member email (AN)
      ];
      
      for (const colName of specificEmailCols) {
        const idx = headers.indexOf(colName);
        if (idx !== -1) {
          emailColIndices.push(idx);
          
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Found specific email column: ${colName} at index ${idx}`, 'DEBUG', 'getClassRecipients');
          }
        }
      }
    }
    
    if (programColIndex === -1 || dayColIndex === -1 || timeColIndex === -1 || emailColIndices.length === 0) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Required columns not found in Daxko sheet. Headers found: ${headers.join(', ')}`, 'ERROR', 'getClassRecipients');
      }
      throw new Error('Required columns not found in Daxko sheet');
    }
    
    // Find students in the class and collect email addresses
    const emails = new Set();
    
    for (let i = 1; i < daxkoData.length; i++) {
      const row = daxkoData[i];
      
      const rowProgram = String(row[programColIndex] || '').trim();
      const rowDay = String(row[dayColIndex] || '').trim();
      const rowTime = String(row[timeColIndex] || '').trim();
      const rowSite = siteColIndex !== -1 ? String(row[siteColIndex] || '').trim() : '';
      
      const programMatch = rowProgram === classInfo.program || 
                           rowProgram.includes(classInfo.program) || 
                           classInfo.program.includes(rowProgram);
      
      const dayMatch = rowDay === classInfo.day || 
                       rowDay.replace('.', '') === classInfo.day.replace('.', '');
      
      const timeMatch = rowTime === classInfo.time || 
                        rowTime.includes(classInfo.time) || 
                        classInfo.time.includes(rowTime);
      
      const siteMatch = siteColIndex === -1 || 
                        classInfo.location === '' || 
                        rowSite === '' || 
                        rowSite === classInfo.location;
      
      if (programMatch && dayMatch && timeMatch && siteMatch) {
        // Add all available email addresses
        for (const emailColIndex of emailColIndices) {
          const email = row[emailColIndex];
          if (email && email.toString().includes('@') && email.toString().includes('.')) {
            emails.add(email.toString().trim());
            
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`Found valid email: ${email} for class: ${classInfo.program}`, 'DEBUG', 'getClassRecipients');
            }
          }
        }
      }
    }
    
    // Convert Set to Array and return
    const emailArray = Array.from(emails);
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Found ${emailArray.length} unique email addresses for class: ${classInfo.program}`, 'INFO', 'getClassRecipients');
    }
    
    return emailArray;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting class recipients: ${error.message}`, 'ERROR', 'getClassRecipients');
    } else {
      Logger.log(`Error getting class recipients: ${error.message}`);
    }
    throw error; // Re-throw to caller
  }
}

/**
 * Sends welcome emails to all participants in group swim lessons.
 * Emails include program information and the swim lesson handbook PDF.
 * 
 * @return {boolean} Success status
 */
function sendWelcomeEmails() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting welcome email send process', 'INFO', 'sendWelcomeEmails');
    }
    
    // Ask for confirmation
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Send Welcome Emails',
        'This will send welcome emails to all participants in group swim lessons who have not already received one. Continue?'
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Send Welcome Emails',
        'This will send welcome emails to all participants in group swim lessons who have not already received one. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Get participants data from Daxko sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName('Daxko');
    
    if (!daxkoSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Daxko sheet not found', 'ERROR', 'sendWelcomeEmails');
      }
      throw new Error('Daxko sheet not found');
    }
    
    // Get the data
    const daxkoData = daxkoSheet.getDataRange().getValues();
    const headers = daxkoData[0];
    
    // Find relevant column indices
    const programColIndex = findColumnIndex(headers, 'Program');
    const sessionColIndex = findColumnIndex(headers, 'Session');
    const firstNameColIndex = findColumnIndex(headers, 'First Name');
    
    // Email columns
    const emailColIndex = findColumnIndex(headers, 'E-mail');
    const primaryEmailColIndex = findColumnIndex(headers, 'Primary Member Email');
    const secondaryEmailColIndex = findColumnIndex(headers, 'Secondary Member Email');
    
    // Welcome Email Sent column
    const welcomeEmailSentColIndex = findColumnIndex(headers, 'Welcome Email Sent');
    
    if (programColIndex === -1 || firstNameColIndex === -1 || 
        (emailColIndex === -1 && primaryEmailColIndex === -1 && secondaryEmailColIndex === -1)) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Required columns not found in Daxko sheet', 'ERROR', 'sendWelcomeEmails');
      }
      throw new Error('Required columns not found in Daxko sheet');
    }
    
    // Prepare to track results
    let totalParticipants = 0;
    let emailsSent = 0;
    let alreadySentCount = 0;
    let noEmailAddressCount = 0;
    let errorCount = 0;
    
    // Get handbook PDF attachment
    const handbookAttachment = getHandbookAttachment();
    
    // Get the HTML email template
    const template = HtmlService.createTemplateFromFile('WelcomeEmail');
    
    // Process each participant
    for (let i = 1; i < daxkoData.length; i++) {
      try {
        const row = daxkoData[i];
        const program = row[programColIndex];
        
        // Skip private lessons and empty rows
        if (!program || program === 'Private Swim Lessons') {
          continue;
        }
        
        totalParticipants++;
        
        // Skip if welcome email already sent
        if (welcomeEmailSentColIndex !== -1 && row[welcomeEmailSentColIndex]) {
          alreadySentCount++;
          continue;
        }
        
        // Get participant details
        const firstName = row[firstNameColIndex] || '';
        const session = sessionColIndex !== -1 ? row[sessionColIndex] || '' : '';
        
        // Collect all available email addresses
        const emails = [];
        if (emailColIndex !== -1 && row[emailColIndex]) {
          emails.push(row[emailColIndex]);
        }
        if (primaryEmailColIndex !== -1 && row[primaryEmailColIndex]) {
          emails.push(row[primaryEmailColIndex]);
        }
        if (secondaryEmailColIndex !== -1 && row[secondaryEmailColIndex]) {
          emails.push(row[secondaryEmailColIndex]);
        }
        
        // Remove duplicates
        const uniqueEmails = [...new Set(emails)];
        
        // Skip if no email address available
        if (uniqueEmails.length === 0) {
          noEmailAddressCount++;
          continue;
        }
        
        // Prepare template data
        template.firstName = firstName;
        template.program = program;
        template.session = session;
        
        // Generate email content
        const htmlContent = template.evaluate().getContent();
        
        // Send email
        GmailApp.sendEmail(
          uniqueEmails.join(','), 
          'Welcome to PenBay YMCA Swim Lessons', 
          'Please view this email with HTML enabled to see the full content.', 
          {
            htmlBody: htmlContent,
            attachments: [handbookAttachment],
            name: "PenBayY - Aquatics",
            from: DEFAULT_SENDER.EMAIL
          }
        );
        
        // Record timestamp in sheet
        if (welcomeEmailSentColIndex !== -1) {
          daxkoSheet.getRange(i + 1, welcomeEmailSentColIndex + 1).setValue(new Date());
        }
        
        emailsSent++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Welcome email sent to ${firstName} (${uniqueEmails.join(', ')})`, 'INFO', 'sendWelcomeEmails');
        }
      } catch (error) {
        errorCount++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error processing row ${i}: ${error.message}`, 'ERROR', 'sendWelcomeEmails');
        } else {
          Logger.log(`Error processing row ${i}: ${error.message}`);
        }
      }
    }
    
    // Log completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Welcome email process completed: ${emailsSent} emails sent, ${alreadySentCount} already sent, ${noEmailAddressCount} missing email, ${errorCount} errors`, 'INFO', 'sendWelcomeEmails');
    }
    
    // Show results
    ui.alert(
      'Welcome Emails Sent',
      `Results:
      - ${emailsSent} welcome emails sent
      - ${alreadySentCount} participants already received emails
      - ${noEmailAddressCount} participants have no email address
      - ${errorCount} errors occurred
      
      Total participants: ${totalParticipants}`,
      ui.ButtonSet.OK
    );
    
    return emailsSent > 0;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'sendWelcomeEmails', 
        'Failed to send welcome emails. Please check your configuration and try again.');
    } else {
      Logger.log(`Error in sendWelcomeEmails: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to send welcome emails: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Gets the handbook PDF attachment for welcome emails.
 * 
 * @return {Blob} The handbook PDF blob
 */
function getHandbookAttachment() {
  try {
    // Get the handbook file ID
    const handbookFileId = '15dyC7yGLX1zy0dDIJ0y8RFjYNBjN-iLL';
    
    // Get the handbook file
    let handbookFile;
    if (typeof GlobalFunctions.safeGetFileById === 'function') {
      handbookFile = GlobalFunctions.safeGetFileById(handbookFileId);
    } else {
      handbookFile = DriveApp.getFileById(handbookFileId);
    }
    
    if (!handbookFile) {
      throw new Error('Could not access handbook file');
    }
    
    // Get the PDF blob
    const pdfBlob = handbookFile.getAs('application/pdf');
    pdfBlob.setName('Swim Lesson Handbook - PenBayY.pdf');
    
    return pdfBlob;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting handbook attachment: ${error.message}`, 'ERROR', 'getHandbookAttachment');
    } else {
      Logger.log(`Error getting handbook attachment: ${error.message}`);
    }
    throw error;
  }
}

/**
 * Helper function to find a column index by name.
 * Falls back to using existing GlobalFunctions if available.
 * 
 * @param {Array} headers - Array of headers
 * @param {string} columnName - Name of the column to find
 * @return {number} Index of the column, or -1 if not found
 */
function findColumnIndex(headers, columnName) {
  // Use GlobalFunctions.findColumnIndex if available
  if (typeof GlobalFunctions.findColumnIndex === 'function') {
    return GlobalFunctions.findColumnIndex(headers, columnName);
  }
  
  // Fallback implementation
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] === columnName) {
      return i;
    }
  }
  
  return -1;
}

/**
 * Tests the welcome email functionality by sending a single test email
 * to the administrator using the first eligible record from the Daxko sheet.
 * 
 * @return {boolean} Success status
 */
function testWelcomeEmail() {
  try {
    const ui = SpreadsheetApp.getUi();
    const adminEmail = "ssullivan@penbayymca.org";
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting welcome email test', 'INFO', 'testWelcomeEmail');
    }
    
    // Ask for confirmation
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Test Welcome Email',
        `This will send a test welcome email to the administrator (${adminEmail}) using data from the first eligible record in the Daxko sheet. Continue?`
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Test Welcome Email',
        `This will send a test welcome email to the administrator (${adminEmail}) using data from the first eligible record in the Daxko sheet. Continue?`,
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Get data from Daxko sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName('Daxko');
    
    if (!daxkoSheet) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Daxko sheet not found', 'ERROR', 'testWelcomeEmail');
      }
      throw new Error('Daxko sheet not found');
    }
    
    // Get the data
    const daxkoData = daxkoSheet.getDataRange().getValues();
    const headers = daxkoData[0];
    
    // Find relevant column indices
    const programColIndex = findColumnIndex(headers, 'Program');
    const sessionColIndex = findColumnIndex(headers, 'Session');
    const firstNameColIndex = findColumnIndex(headers, 'First Name');
    
    if (programColIndex === -1 || firstNameColIndex === -1) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Required columns not found in Daxko sheet', 'ERROR', 'testWelcomeEmail');
      }
      throw new Error('Required columns not found in Daxko sheet');
    }
    
    // Find the first eligible record
    let testRecord = null;
    let recordIndex = 0;
    
    for (let i = 1; i < daxkoData.length; i++) {
      const row = daxkoData[i];
      const program = row[programColIndex];
      
      // Look for non-private lesson with a program name
      if (program && program !== 'Private Swim Lessons') {
        testRecord = row;
        recordIndex = i;
        break;
      }
    }
    
    if (!testRecord) {
      ui.alert('Test Failed', 'No eligible records found in the Daxko sheet.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get handbook attachment
    const handbookAttachment = getHandbookAttachment();
    
    // Get record data
    const firstName = testRecord[firstNameColIndex] || 'Student';
    const program = testRecord[programColIndex] || 'Swim Lessons';
    const session = sessionColIndex !== -1 ? testRecord[sessionColIndex] || 'Current' : 'Current';
    
    // Get the HTML email template
    const template = HtmlService.createTemplateFromFile('WelcomeEmail');
    
    // Set template data
    template.firstName = firstName;
    template.program = program;
    template.session = session;
    
    // Generate email content
    const htmlContent = template.evaluate().getContent();
    
    // Send test email to administrator
    GmailApp.sendEmail(
      adminEmail, 
      '[TEST] Welcome to PenBay YMCA Swim Lessons', 
      'Please view this email with HTML enabled to see the full content.', 
      {
        htmlBody: htmlContent,
        attachments: [handbookAttachment],
        name: "PenBayY - Aquatics [TEST]",
        from: DEFAULT_SENDER.EMAIL
      }
    );
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Test welcome email sent to ${adminEmail} using data for ${firstName}`, 'INFO', 'testWelcomeEmail');
    }
    
    // Show success message
    ui.alert(
      'Test Email Sent',
      `A test welcome email has been sent to ${adminEmail} using data from record #${recordIndex + 1}:\n\n` +
      `First Name: ${firstName}\n` +
      `Program: ${program}\n` +
      `Session: ${session}`,
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'testWelcomeEmail', 
        'Failed to send test welcome email. Please check your configuration and try again.');
    } else {
      Logger.log(`Error in testWelcomeEmail: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to send test welcome email: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

const CommunicationModule = {
  emailClassParticipants: emailClassParticipants,
  sendClassAnnouncements: sendClassAnnouncements,
  sendReadyAnnouncements: sendReadyAnnouncements,
  sendWelcomeEmails: sendWelcomeEmails,
  testWelcomeEmail: testWelcomeEmail
};