/**
 * YSL Hub v2 Reporting Module
 * 
 * This module handles the generation and distribution of student progress reports,
 * including mid-session and end-session evaluations.
 * 
 * @author PenBay YMCA
 * @version 2.0
 * @date 2025-04-27
 */

// Global constants
const SENDER_EMAIL = "ssullivan@penbayymca.org";
const SUMMARY_RECIPIENT = "ssullivan@penbayymca.org";

/**
 * Generates and sends mid-session reports for all selected classes.
 * 
 * @return {boolean} Success status
 */
function generateMidSessionReports() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting mid-session report generation', 'INFO', 'generateMidSessionReports');
    }
    
    // Ask for confirmation
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Generate Mid-Session Reports',
        'This will generate and send mid-session reports for all selected classes. Continue?'
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Generate Mid-Session Reports',
        'This will generate and send mid-session reports for all selected classes. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Get configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    // Validate report template folder
    let templateFolderId;
    if (typeof GlobalFunctions.extractIdFromUrl === 'function') {
      templateFolderId = GlobalFunctions.extractIdFromUrl(config.reportTemplateUrl);
    } else {
      // Fallback to basic extraction
      templateFolderId = extractIdFromUrl(config.reportTemplateUrl);
    }
    
    if (!templateFolderId) {
      ui.alert('Error', 'Invalid report template folder URL in system configuration.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get folder safely
    let templateFolder;
    if (typeof GlobalFunctions.safeGetFolderById === 'function') {
      templateFolder = GlobalFunctions.safeGetFolderById(templateFolderId);
    } else {
      // Fallback to direct access
      try {
        templateFolder = DriveApp.getFolderById(templateFolderId);
      } catch (error) {
        ui.alert('Error', `Could not access report template folder: ${error.message}`, ui.ButtonSet.OK);
        return false;
      }
    }
    
    if (!templateFolder) {
      ui.alert('Error', 'Could not access report template folder. Please check permissions and try again.', ui.ButtonSet.OK);
      return false;
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
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Processing ${selectedClasses.length} classes for mid-session reports`, 'INFO', 'generateMidSessionReports');
    }
    
    // Process each selected class
    let totalReports = 0;
    let reportsSent = 0;
    let errorCount = 0;
    const summary = [];
    
    // Process each class
    for (const classInfo of selectedClasses) {
      try {
        // Get class sheet
        const sheetName = `Class_${classInfo.program.replace(/[^a-zA-Z0-9]/g, '')}_${classInfo.day.replace(/[^a-zA-Z0-9]/g, '')}_${classInfo.time.replace(/[^a-zA-Z0-9]/g, '')}`;
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const classSheet = ss.getSheetByName(sheetName);
        
        if (!classSheet) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Sheet not found for class: ${classInfo.program} (${classInfo.day}, ${classInfo.time})`, 'WARNING', 'generateMidSessionReports');
          }
          continue;
        }
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Processing class: ${classInfo.program} (${classInfo.day}, ${classInfo.time})`, 'INFO', 'generateMidSessionReports');
        }
        
        // Get class data
        const classData = classSheet.getDataRange().getValues();
        const headers = classData[0];
        
        // Find relevant columns
        const nameColIndex = headers.indexOf('Swimmer');
        const dobColIndex = headers.indexOf('DOB');
        const emailColsIndices = [];
        
        // Find email columns
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] && headers[i].toString().toLowerCase().includes('email')) {
            emailColsIndices.push(i);
          }
        }
        
        if (nameColIndex === -1 || emailColsIndices.length === 0) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Required columns not found in sheet for class: ${classInfo.program}`, 'WARNING', 'generateMidSessionReports');
          }
          continue;
        }
        
        // Get template
        const templateName = `MidSession Report Template - ${classInfo.program}`;
        const templateFile = findReportTemplate(templateFolderId, templateName);
        
        if (!templateFile) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Template not found: ${templateName}`, 'WARNING', 'generateMidSessionReports');
          }
          continue;
        }
        
        // Process students
        for (let i = 1; i < classData.length; i++) {
          const student = classData[i];
          const studentName = student[nameColIndex];
          
          if (!studentName) continue;
          
          // Find email
          let email = null;
          for (const emailColIndex of emailColsIndices) {
            if (student[emailColIndex]) {
              email = student[emailColIndex];
              break;
            }
          }
          
          if (!email) {
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`No email found for student: ${studentName}`, 'WARNING', 'generateMidSessionReports');
            }
            continue;
          }
          
          // Create report data
          const reportData = {};
          headers.forEach((header, index) => {
            if (header) {
              reportData[`{{${header}}}`] = student[index] || '';
            }
          });
          
          // Add static fields
          reportData['{{Session}}'] = config.sessionName || '';
          reportData['{{Date}}'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');
          reportData['{{Instructor}}'] = classInfo.instructor || '';
          
          // Generate PDF
          const pdfBlob = createReportPDF(templateFile, studentName, reportData, 'MidSession');
          
          // Send email with PDF
          if (pdfBlob) {
            sendReportEmail(email, studentName, classInfo.instructor, pdfBlob, 'Mid-Session');
            
            // Update summary
            summary.push(`✅ ${studentName} (${email}) – ${classInfo.program}`);
            reportsSent++;
            
            // Update tracking in sheet if sent-status column exists
            const sentStatusColIndex = headers.indexOf('Mid Sent');
            if (sentStatusColIndex !== -1) {
              classSheet.getRange(i + 1, sentStatusColIndex + 1).setValue(new Date());
            }
            
            totalReports++;
            
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`Mid-session report sent for: ${studentName}`, 'INFO', 'generateMidSessionReports');
            }
          }
        }
      } catch (error) {
        errorCount++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error processing class ${classInfo.program}: ${error.message}`, 'ERROR', 'generateMidSessionReports');
        } else {
          Logger.log(`Error processing class ${classInfo.program}: ${error.message}`);
        }
      }
    }
    
    // Send summary email to admin
    if (summary.length > 0) {
      sendSummaryEmail(summary, reportsSent, 'Mid-Session');
    }
    
    // Log completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Mid-session report generation completed: ${reportsSent} reports sent, ${errorCount} errors`, 'INFO', 'generateMidSessionReports');
    }
    
    // Show results
    ui.alert(
      'Report Generation Complete',
      `${reportsSent} of ${totalReports} mid-session reports were sent successfully.\n` +
      (errorCount > 0 ? `${errorCount} errors occurred. Check logs for details.` : ''),
      ui.ButtonSet.OK
    );
    
    return reportsSent > 0;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'generateMidSessionReports', 
        'Failed to generate mid-session reports. Please check your configuration and try again.');
    } else {
      Logger.log(`Mid-session report generation error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to generate mid-session reports: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Generates and sends end-session reports for all selected classes.
 * 
 * @return {boolean} Success status
 */
function generateEndSessionReports() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting end-session report generation', 'INFO', 'generateEndSessionReports');
    }
    
    // Ask for confirmation
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Generate End-Session Reports',
        'This will generate and send end-session reports for all selected classes. Continue?'
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Generate End-Session Reports',
        'This will generate and send end-session reports for all selected classes. Continue?',
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Get configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    // Validate report template folder
    let templateFolderId;
    if (typeof GlobalFunctions.extractIdFromUrl === 'function') {
      templateFolderId = GlobalFunctions.extractIdFromUrl(config.reportTemplateUrl);
    } else {
      // Fallback to basic extraction
      templateFolderId = extractIdFromUrl(config.reportTemplateUrl);
    }
    
    if (!templateFolderId) {
      ui.alert('Error', 'Invalid report template folder URL in system configuration.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get folder safely
    let templateFolder;
    if (typeof GlobalFunctions.safeGetFolderById === 'function') {
      templateFolder = GlobalFunctions.safeGetFolderById(templateFolderId);
    } else {
      // Fallback to direct access
      try {
        templateFolder = DriveApp.getFolderById(templateFolderId);
      } catch (error) {
        ui.alert('Error', `Could not access report template folder: ${error.message}`, ui.ButtonSet.OK);
        return false;
      }
    }
    
    if (!templateFolder) {
      ui.alert('Error', 'Could not access report template folder. Please check permissions and try again.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get end-session folder
    let endSessionFolder = null;
    const subfolders = templateFolder.getFolders();
    
    while (subfolders.hasNext()) {
      const folder = subfolders.next();
      if (folder.getName().includes('End-Session')) {
        endSessionFolder = folder;
        break;
      }
    }
    
    if (!endSessionFolder) {
      // Try to create the folder
      try {
        endSessionFolder = templateFolder.createFolder('End-Session Reports');
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage('Created new End-Session Reports folder', 'INFO', 'generateEndSessionReports');
        }
      } catch (error) {
        ui.alert('Error', 'End-Session Reports folder not found and could not be created.', ui.ButtonSet.OK);
        return false;
      }
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
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Processing ${selectedClasses.length} classes for end-session reports`, 'INFO', 'generateEndSessionReports');
    }
    
    // Process each selected class
    let totalReports = 0;
    let reportsSent = 0;
    let errorCount = 0;
    const summary = [];
    
    // Process each class
    for (const classInfo of selectedClasses) {
      try {
        // Get class sheet
        const sheetName = `Class_${classInfo.program.replace(/[^a-zA-Z0-9]/g, '')}_${classInfo.day.replace(/[^a-zA-Z0-9]/g, '')}_${classInfo.time.replace(/[^a-zA-Z0-9]/g, '')}`;
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const classSheet = ss.getSheetByName(sheetName);
        
        if (!classSheet) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Sheet not found for class: ${classInfo.program} (${classInfo.day}, ${classInfo.time})`, 'WARNING', 'generateEndSessionReports');
          }
          continue;
        }
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Processing class: ${classInfo.program} (${classInfo.day}, ${classInfo.time})`, 'INFO', 'generateEndSessionReports');
        }
        
        // Get class data
        const classData = classSheet.getDataRange().getValues();
        const headers = classData[0];
        
        // Find relevant columns
        const nameColIndex = headers.indexOf('Swimmer');
        const dobColIndex = headers.indexOf('DOB');
        const emailColsIndices = [];
        
        // Find email columns
        for (let i = 0; i < headers.length; i++) {
          if (headers[i] && headers[i].toString().toLowerCase().includes('email')) {
            emailColsIndices.push(i);
          }
        }
        
        if (nameColIndex === -1 || emailColsIndices.length === 0) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Required columns not found in sheet for class: ${classInfo.program}`, 'WARNING', 'generateEndSessionReports');
          }
          continue;
        }
        
        // Get template
        const templateName = `EndSession Report Template - ${classInfo.program}`;
        const endSessionFolderId = endSessionFolder ? endSessionFolder.getId() : null;
        if (!endSessionFolderId) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage('End-Session folder ID not available', 'ERROR', 'generateEndSessionReports');
          }
          continue;
        }
        
        const templateFile = findReportTemplate(endSessionFolderId, templateName);
        
        if (!templateFile) {
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Template not found: ${templateName}`, 'WARNING', 'generateEndSessionReports');
          }
          continue;
        }
        
        // Process students
        for (let i = 1; i < classData.length; i++) {
          const student = classData[i];
          const studentName = student[nameColIndex];
          
          if (!studentName) continue;
          
          // Find email
          let email = null;
          for (const emailColIndex of emailColsIndices) {
            if (student[emailColIndex]) {
              email = student[emailColIndex];
              break;
            }
          }
          
          if (!email) {
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`No email found for student: ${studentName}`, 'WARNING', 'generateEndSessionReports');
            }
            continue;
          }
          
          // Create report data
          const reportData = {};
          headers.forEach((header, index) => {
            if (header) {
              reportData[`{{${header}}}`] = student[index] || '';
            }
          });
          
          // Add static fields
          reportData['{{Session}}'] = config.sessionName || '';
          reportData['{{Date}}'] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM dd, yyyy');
          reportData['{{Instructor}}'] = classInfo.instructor || '';
          
          // Generate PDF
          const pdfBlob = createReportPDF(templateFile, studentName, reportData, 'End Session');
          
          // Send email with PDF
          if (pdfBlob) {
            sendReportEmail(email, studentName, classInfo.instructor, pdfBlob, 'End-Session');
            
            // Update summary
            summary.push(`✅ ${studentName} (${email}) – ${classInfo.program}`);
            reportsSent++;
            
            // Update tracking in sheet if sent-status column exists
            const sentStatusColIndex = headers.indexOf('End Sent');
            if (sentStatusColIndex !== -1) {
              classSheet.getRange(i + 1, sentStatusColIndex + 1).setValue(new Date());
            }
            
            totalReports++;
            
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`End-session report sent for: ${studentName}`, 'INFO', 'generateEndSessionReports');
            }
          }
        }
      } catch (error) {
        errorCount++;
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error processing class ${classInfo.program}: ${error.message}`, 'ERROR', 'generateEndSessionReports');
        } else {
          Logger.log(`Error processing class ${classInfo.program}: ${error.message}`);
        }
      }
    }
    
    // Send summary email to admin
    if (summary.length > 0) {
      sendSummaryEmail(summary, reportsSent, 'End-Session');
    }
    
    // Log completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`End-session report generation completed: ${reportsSent} reports sent, ${errorCount} errors`, 'INFO', 'generateEndSessionReports');
    }
    
    // Show results
    ui.alert(
      'Report Generation Complete',
      `${reportsSent} of ${totalReports} end-session reports were sent successfully.\n` +
      (errorCount > 0 ? `${errorCount} errors occurred. Check logs for details.` : ''),
      ui.ButtonSet.OK
    );
    
    return reportsSent > 0;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'generateEndSessionReports', 
        'Failed to generate end-session reports. Please check your configuration and try again.');
    } else {
      Logger.log(`End-session report generation error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', `Failed to generate end-session reports: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
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
 * Finds a report template by name in the specified folder.
 * 
 * @param {string} folderId - ID of the folder containing report templates
 * @param {string} templateName - Name of the template to find
 * @return {Object|null} The template file, or null if not found
 */
function findReportTemplate(folderId, templateName) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Looking for template: ${templateName} in folder: ${folderId}`, 'INFO', 'findReportTemplate');
    }
    
    // Safely get folder
    let folder;
    if (typeof GlobalFunctions.safeGetFolderById === 'function') {
      folder = GlobalFunctions.safeGetFolderById(folderId);
    } else {
      // Fallback to direct access
      try {
        folder = DriveApp.getFolderById(folderId);
      } catch (error) {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Could not access folder with ID: ${folderId}`, 'ERROR', 'findReportTemplate');
        }
        return null;
      }
    }
    
    if (!folder) {
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Folder not found with ID: ${folderId}`, 'ERROR', 'findReportTemplate');
      }
      return null;
    }
    
    // First try exact match
    const files = folder.getFilesByName(templateName);
    
    if (files.hasNext()) {
      const file = files.next();
      
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(`Found exact match template: ${file.getName()}`, 'INFO', 'findReportTemplate');
      }
      
      return file;
    }
    
    // If exact match not found, try partial match
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`No exact match found, trying partial match for: ${templateName}`, 'INFO', 'findReportTemplate');
    }
    
    const allFiles = folder.getFiles();
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      if (file.getName().includes(templateName)) {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Found partial match template: ${file.getName()}`, 'INFO', 'findReportTemplate');
        }
        
        return file;
      }
    }
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`No template found for: ${templateName}`, 'WARNING', 'findReportTemplate');
    }
    
    return null;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error finding template: ${error.message}`, 'ERROR', 'findReportTemplate');
    } else {
      Logger.log(`Error finding template: ${error.message}`);
    }
    return null;
  }
}

/**
 * Creates a PDF report from a template.
 * 
 * @param {Object} templateFile - The template file
 * @param {string} studentName - Name of the student
 * @param {Object} reportData - The data to merge into the template
 * @param {string} reportType - Type of report (MidSession or End Session)
 * @return {Object} PDF blob
 */
function createReportPDF(templateFile, studentName, reportData, reportType) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Creating ${reportType} PDF for: ${studentName}`, 'INFO', 'createReportPDF');
    }
    
    const safeStudentName = studentName.replace(/[^\w\s]/gi, '');
    const copiedFile = templateFile.makeCopy(`${reportType} Report - ${safeStudentName}`);
    
    // Allow time for the copy to process
    Utilities.sleep(1000);
    
    const slideDeck = SlidesApp.openById(copiedFile.getId());
    const slides = slideDeck.getSlides();
    
    // Process each slide
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      
      // Replace merge tags in all shapes with text
      slide.getShapes().forEach(shape => {
        if (shape.getText) {
          const textRange = shape.getText();
          if (textRange) {
            // Replace all template tags
            for (const [tag, value] of Object.entries(reportData)) {
              if (tag && value !== undefined && value !== null) {
                textRange.replaceAllText(tag, value.toString());
              }
            }
          }
        }
      });
      
      // Check for tables on the slide
      const tables = slide.getTables();
      for (const table of tables) {
        const numRows = table.getNumRows();
        const numCols = table.getNumColumns();
        
        // Process each cell in the table
        for (let row = 0; row < numRows; row++) {
          for (let col = 0; col < numCols; col++) {
            const cell = table.getCell(row, col);
            const text = cell.getText();
            
            // Replace all template tags
            for (const [tag, value] of Object.entries(reportData)) {
              if (tag && value !== undefined && value !== null) {
                text.replaceAllText(tag, value.toString());
              }
            }
          }
        }
      }
    }
    
    slideDeck.saveAndClose();
    
    // Export to PDF
    const pdfBlob = DriveApp.getFileById(copiedFile.getId())
      .getAs("application/pdf")
      .setName(`${reportType} Report - ${safeStudentName}.pdf`);
    
    // Clean up the copied file
    DriveApp.getFileById(copiedFile.getId()).setTrashed(true);
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Successfully created PDF for: ${studentName}`, 'INFO', 'createReportPDF');
    }
    
    return pdfBlob;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error creating PDF for ${studentName}: ${error.message}`, 'ERROR', 'createReportPDF');
    } else {
      Logger.log(`Error creating PDF for ${studentName}: ${error.message}`);
    }
    return null;
  }
}

/**
 * Sends an email with a report attachment.
 * 
 * @param {string} email - Recipient email
 * @param {string} studentName - Name of the student
 * @param {string} instructor - Name of the instructor
 * @param {Object} pdfBlob - The PDF attachment
 * @param {string} reportType - Type of report (Mid-Session or End-Session)
 */
function sendReportEmail(email, studentName, instructor, pdfBlob, reportType) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Sending ${reportType} report email to: ${email} for ${studentName}`, 'INFO', 'sendReportEmail');
    }
    
    // Different subject lines for different report types
    const emailSubject = reportType === "Mid-Session" 
      ? `Progress Update: ${studentName}'s Mid-Session Swim Report` 
      : `Achievement Report: ${studentName}'s Swim Assessment`;
    
    // Generate email body
    let emailBody;
    if (reportType === "Mid-Session") {
      emailBody = 
        `<p>Hello,</p>
        <p>Attached please find ${studentName}'s mid-session swim progress report. This report provides feedback on skills that have been mastered and areas that are still being developed.</p>
        <p>If you have any questions about your child's progress, please feel free to speak with your instructor, ${instructor || 'your instructor'}, or contact the aquatics office.</p>
        <p>Thank you for your participation in our swim lesson program!</p>
        <p>Best regards,<br>
        PenBay YMCA Aquatics</p>`;
    } else {
      emailBody = 
        `<p>Hello,</p>
        <p>Attached please find ${studentName}'s end-session swim assessment. This report summarizes your child's progress throughout the session and provides recommendations for next steps.</p>
        <p>We've also included information about our upcoming swim session for your convenience.</p>
        <p>If you have any questions about this assessment or future class recommendations, please contact the aquatics office.</p>
        <p>Thank you for your participation in our swim lesson program!</p>
        <p>Best regards,<br>
        PenBay YMCA Aquatics</p>`;
    }
    
    // Prepare attachments
    let attachments = [pdfBlob];
    
    // Add Parent Handbook only to end-session emails if available
    if (reportType === "End-Session") {
      try {
        const config = AdministrativeModule.getSystemConfiguration();
        if (config.parentHandbookUrl) {
          const handbookId = GlobalFunctions.extractIdFromUrl(config.parentHandbookUrl);
          if (handbookId) {
            try {
              const handbookFile = GlobalFunctions.safeGetFileById(handbookId);
              
              if (handbookFile) {
                const handbookPdf = handbookFile
                  .getAs("application/pdf")
                  .setName("PenBay YMCA Swim Lessons Handbook.pdf");
                
                attachments.push(handbookPdf);
                
                if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
                  ErrorHandling.logMessage('Added parent handbook to email attachments', 'INFO', 'sendReportEmail');
                }
              }
            } catch (e) {
              if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
                ErrorHandling.logMessage(`Error accessing parent handbook: ${e.message}`, 'WARNING', 'sendReportEmail');
              } else {
                Logger.log(`Error accessing parent handbook: ${e.message}`);
              }
            }
          }
        }
      } catch (error) {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error attaching parent handbook: ${error.message}`, 'WARNING', 'sendReportEmail');
        } else {
          Logger.log(`Error attaching parent handbook: ${error.message}`);
        }
      }
    }
    
    // Send the email with appropriate attachments
    GmailApp.sendEmail(email, emailSubject, '', {
      htmlBody: emailBody,
      attachments: attachments,
      from: SENDER_EMAIL,
      name: "PenBayY - Aquatics"
    });
    
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Successfully sent ${reportType} report email to: ${email}`, 'INFO', 'sendReportEmail');
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error sending email to ${email}: ${error.message}`, 'ERROR', 'sendReportEmail');
    } else {
      Logger.log(`Error sending email to ${email}: ${error.message}`);
    }
  }
}

/**
 * Sends a summary email to the administrator.
 * 
 * @param {Array} summary - Summary of sent reports
 * @param {number} totalSent - Total number of reports sent
 * @param {string} reportType - Type of report (Mid-Session or End-Session)
 */
function sendSummaryEmail(summary, totalSent, reportType) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Sending ${reportType} summary email with ${totalSent} reports`, 'INFO', 'sendSummaryEmail');
    }
    
    const subject = `${reportType} Reports Sent: ${totalSent} ${totalSent === 1 ? 'swimmer' : 'swimmers'}`;
    const body = totalSent > 0
      ? `The following reports were sent:\n\n${summary.join("\n")}`
      : `No reports were sent. Please check for missing data or templates.`;
  
    GmailApp.sendEmail(SUMMARY_RECIPIENT, subject, body, {
      from: SENDER_EMAIL,
      name: `${reportType} Report Bot`
    });
  
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`${reportType} summary email sent to ${SUMMARY_RECIPIENT}`, 'INFO', 'sendSummaryEmail');
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error sending summary email: ${error.message}`, 'ERROR', 'sendSummaryEmail');
    } else {
      Logger.log(`Error sending summary email: ${error.message}`);
    }
  }
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

// Make functions available to other modules
const ReportingModule = {
  generateMidSessionReports: generateMidSessionReports,
  generateEndSessionReports: generateEndSessionReports
};