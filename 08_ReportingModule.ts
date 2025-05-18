/**
 * YSL Hub v2 Reporting Module
 * 
 * This module handles the generation of reports for swim lessons, including
 * mid-session progress reports and end-session assessment reports. It provides
 * functions for creating, customizing, and distributing reports to parents.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

// Report types
const REPORT_TYPES = {
  MID_SESSION: 'Mid-Session Progress Report',
  END_SESSION: 'End-Session Assessment Report'
};

// Placeholder patterns for report templates
const PLACEHOLDERS = {
  STUDENT_NAME: '{{student_name}}',
  STUDENT_ID: '{{student_id}}',
  CLASS_NAME: '{{class_name}}',
  LEVEL: '{{level}}',
  INSTRUCTOR_NAME: '{{instructor_name}}',
  SESSION_NAME: '{{session_name}}',
  REPORT_DATE: '{{report_date}}',
  PARENT_NAME: '{{parent_name}}',
  SKILLS_TABLE: '{{skills_table}}',
  COMMENTS: '{{comments}}',
  NEXT_STEPS: '{{next_steps}}',
  ASSESSMENT_SUMMARY: '{{assessment_summary}}',
  NEXT_LEVEL: '{{next_level}}',
  RECOMMENDED_CLASSES: '{{recommended_classes}}'
};

/**
 * Generates mid-session progress reports for selected classes
 * 
 * @returns Success status
 */
function generateMidSessionReports() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Generating mid-session progress reports', 'INFO', 'generateMidSessionReports');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Get template information
    const reportTemplateUrl = getReportTemplateUrl(REPORT_TYPES.MID_SESSION);
    
    if (!reportTemplateUrl) {
      ui.alert(
        'Missing Template',
        'The report template URL is not configured or no mid-session template was found. Please update system configuration first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Show class selection UI
    const selectedClasses = showClassSelectionUI('Select Classes for Mid-Session Reports');
    
    if (!selectedClasses || selectedClasses.length === 0) {
      return true; // User cancelled or no classes selected, but not an error
    }
    
    // Check for Group Lesson Tracker data
    const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Group Lesson Tracker');
    
    if (!trackerSheet) {
      ui.alert(
        'Missing Data',
        'The Group Lesson Tracker sheet is not found. Please generate it first and enter assessment data.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Ask for report folder destination
    const folderResult = ui.prompt(
      'Report Destination',
      'Enter the name for the reports folder (or leave blank to use default):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (folderResult.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    // Set up folder
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || 'Current Session';
    const folderName = folderResult.getResponseText().trim() || `YSL ${sessionName} Mid-Session Reports`;
    
    // Create or get folder
    let reportFolder;
    try {
      reportFolder = createReportFolder(folderName);
      
      if (!reportFolder) {
        ui.alert(
          'Folder Error',
          'Failed to create or access the reports folder. Please check permissions and try again.',
          ui.ButtonSet.OK
        );
        return false;
      }
    } catch (folderError) {
      if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
        ErrorHandling.handleError(folderError, 'generateMidSessionReports', 
          'Error creating report folder. Please check drive permissions and try again.');
      } else {
        Logger.log(`Error creating report folder: ${folderError.message}`);
        ui.alert(
          'Folder Error',
          `Failed to create report folder: ${folderError.message}`,
          ui.ButtonSet.OK
        );
      }
      return false;
    }
    
    // Get template document
    let templateDoc;
    try {
      templateDoc = DocumentApp.openByUrl(reportTemplateUrl);
    } catch (docError) {
      ui.alert(
        'Template Error',
        `Failed to open the report template. Error: ${docError.message}\n\nPlease check that the template URL is correct and you have permission to access it.`,
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Process each selected class
    let totalReports = 0;
    let totalFailures = 0;
    const reportLinks = [];
    const errors = [];
    
    for (const classId of selectedClasses) {
      const classDetails = getClassDetails(classId);
      
      if (!classDetails) {
        errors.push(`Could not get details for class ID ${classId}`);
        continue;
      }
      
      // Get students in this class with assessment data
      const students = getStudentsWithAssessments(classId);
      
      if (students.length === 0) {
        errors.push(`No students with assessment data found for class ${classDetails.className}`);
        continue;
      }
      
      // Create a class folder
      const classFolderName = `${classDetails.className} - ${classDetails.day} ${classDetails.time}`;
      let classFolder;
      try {
        classFolder = createSubfolder(reportFolder, classFolderName);
      } catch (folderError) {
        errors.push(`Failed to create folder for class ${classDetails.className}: ${folderError.message}`);
        continue;
      }
      
      // Generate reports for each student
      for (const student of students) {
        try {
          // Create a copy of the template
          const reportName = `${REPORT_TYPES.MID_SESSION} - ${student.firstName} ${student.lastName}`;
          const reportDoc = templateDoc.makeCopy(reportName, classFolder);
          
          // Open the copy for editing
          const doc = DocumentApp.openById(reportDoc.getId());
          const body = doc.getBody();
          
          // Replace placeholders with student data
          replacePlaceholders(body, student, classDetails, sessionName);
          
          // Save and close
          doc.saveAndClose();
          
          // Add to report links
          reportLinks.push({
            student: `${student.firstName} ${student.lastName}`,
            url: reportDoc.getUrl(),
            class: classDetails.className
          });
          
          totalReports++;
        } catch (reportError) {
          totalFailures++;
          errors.push(`Error creating report for ${student.firstName} ${student.lastName}: ${reportError.message}`);
          
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Error creating report for ${student.firstName} ${student.lastName}: ${reportError.message}`, 'ERROR', 'generateMidSessionReports');
          }
        }
      }
    }
    
    // Create a summary sheet
    createReportSummarySheet(
      reportLinks, 
      errors, 
      `${REPORT_TYPES.MID_SESSION} - Summary`, 
      reportFolder.getUrl()
    );
    
    // Show results
    if (totalFailures > 0) {
      ui.alert(
        'Reports Generated with Errors',
        `Successfully created ${totalReports} reports with ${totalFailures} failures.\n\nThe reports folder is available here: ${reportFolder.getUrl()}\n\nA summary sheet has been created with links to all reports.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Reports Generated',
        `Successfully created ${totalReports} reports.\n\nThe reports folder is available here: ${reportFolder.getUrl()}\n\nA summary sheet has been created with links to all reports.`,
        ui.ButtonSet.OK
      );
    }
    
    // Ask if user wants to send emails
    const emailResult = ui.alert(
      'Send Email Reports',
      'Would you like to send these reports to parents via email?',
      ui.ButtonSet.YES_NO
    );
    
    if (emailResult === ui.Button.YES) {
      return emailReports(reportLinks, REPORT_TYPES.MID_SESSION);
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'generateMidSessionReports', 
        'Error generating mid-session reports. Please try again or contact support.');
    } else {
      Logger.log(`Error generating mid-session reports: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to generate mid-session reports: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Generates end-session assessment reports for selected classes
 * 
 * @returns Success status
 */
function generateEndSessionReports() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Generating end-session assessment reports', 'INFO', 'generateEndSessionReports');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Get template information
    const reportTemplateUrl = getReportTemplateUrl(REPORT_TYPES.END_SESSION);
    
    if (!reportTemplateUrl) {
      ui.alert(
        'Missing Template',
        'The report template URL is not configured or no end-session template was found. Please update system configuration first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Show class selection UI
    const selectedClasses = showClassSelectionUI('Select Classes for End-Session Reports');
    
    if (!selectedClasses || selectedClasses.length === 0) {
      return true; // User cancelled or no classes selected, but not an error
    }
    
    // Check for Group Lesson Tracker data
    const trackerSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Group Lesson Tracker');
    
    if (!trackerSheet) {
      ui.alert(
        'Missing Data',
        'The Group Lesson Tracker sheet is not found. Please generate it first and enter assessment data.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Ask for report folder destination
    const folderResult = ui.prompt(
      'Report Destination',
      'Enter the name for the reports folder (or leave blank to use default):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (folderResult.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    // Set up folder
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || 'Current Session';
    const folderName = folderResult.getResponseText().trim() || `YSL ${sessionName} End-Session Reports`;
    
    // Create or get folder
    let reportFolder;
    try {
      reportFolder = createReportFolder(folderName);
      
      if (!reportFolder) {
        ui.alert(
          'Folder Error',
          'Failed to create or access the reports folder. Please check permissions and try again.',
          ui.ButtonSet.OK
        );
        return false;
      }
    } catch (folderError) {
      if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
        ErrorHandling.handleError(folderError, 'generateEndSessionReports', 
          'Error creating report folder. Please check drive permissions and try again.');
      } else {
        Logger.log(`Error creating report folder: ${folderError.message}`);
        ui.alert(
          'Folder Error',
          `Failed to create report folder: ${folderError.message}`,
          ui.ButtonSet.OK
        );
      }
      return false;
    }
    
    // Get template document
    let templateDoc;
    try {
      templateDoc = DocumentApp.openByUrl(reportTemplateUrl);
    } catch (docError) {
      ui.alert(
        'Template Error',
        `Failed to open the report template. Error: ${docError.message}\n\nPlease check that the template URL is correct and you have permission to access it.`,
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Process each selected class
    let totalReports = 0;
    let totalFailures = 0;
    const reportLinks = [];
    const errors = [];
    
    for (const classId of selectedClasses) {
      const classDetails = getClassDetails(classId);
      
      if (!classDetails) {
        errors.push(`Could not get details for class ID ${classId}`);
        continue;
      }
      
      // Get students in this class with assessment data
      const students = getStudentsWithAssessments(classId);
      
      if (students.length === 0) {
        errors.push(`No students with assessment data found for class ${classDetails.className}`);
        continue;
      }
      
      // Create a class folder
      const classFolderName = `${classDetails.className} - ${classDetails.day} ${classDetails.time}`;
      let classFolder;
      try {
        classFolder = createSubfolder(reportFolder, classFolderName);
      } catch (folderError) {
        errors.push(`Failed to create folder for class ${classDetails.className}: ${folderError.message}`);
        continue;
      }
      
      // Analyze end-session results for each student
      for (const student of students) {
        try {
          // Calculate next level recommendation
          const nextLevel = calculateNextLevel(student);
          
          // Create a copy of the template
          const reportName = `${REPORT_TYPES.END_SESSION} - ${student.firstName} ${student.lastName}`;
          const reportDoc = templateDoc.makeCopy(reportName, classFolder);
          
          // Open the copy for editing
          const doc = DocumentApp.openById(reportDoc.getId());
          const body = doc.getBody();
          
          // Replace placeholders with student data
          replacePlaceholders(body, student, classDetails, sessionName, nextLevel);
          
          // Add assessment summary
          const summary = generateAssessmentSummary(student);
          body.replaceText(PLACEHOLDERS.ASSESSMENT_SUMMARY, summary);
          
          // Add recommended classes for next session
          const recommendations = generateRecommendations(student, nextLevel);
          body.replaceText(PLACEHOLDERS.RECOMMENDED_CLASSES, recommendations);
          
          // Save and close
          doc.saveAndClose();
          
          // Add to report links
          reportLinks.push({
            student: `${student.firstName} ${student.lastName}`,
            url: reportDoc.getUrl(),
            class: classDetails.className,
            nextLevel: nextLevel
          });
          
          totalReports++;
        } catch (reportError) {
          totalFailures++;
          errors.push(`Error creating report for ${student.firstName} ${student.lastName}: ${reportError.message}`);
          
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Error creating report for ${student.firstName} ${student.lastName}: ${reportError.message}`, 'ERROR', 'generateEndSessionReports');
          }
        }
      }
    }
    
    // Create a summary sheet
    createReportSummarySheet(
      reportLinks, 
      errors, 
      `${REPORT_TYPES.END_SESSION} - Summary`, 
      reportFolder.getUrl()
    );
    
    // Show results
    if (totalFailures > 0) {
      ui.alert(
        'Reports Generated with Errors',
        `Successfully created ${totalReports} reports with ${totalFailures} failures.\n\nThe reports folder is available here: ${reportFolder.getUrl()}\n\nA summary sheet has been created with links to all reports.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Reports Generated',
        `Successfully created ${totalReports} reports.\n\nThe reports folder is available here: ${reportFolder.getUrl()}\n\nA summary sheet has been created with links to all reports.`,
        ui.ButtonSet.OK
      );
    }
    
    // Ask if user wants to send emails
    const emailResult = ui.alert(
      'Send Email Reports',
      'Would you like to send these reports to parents via email?',
      ui.ButtonSet.YES_NO
    );
    
    if (emailResult === ui.Button.YES) {
      return emailReports(reportLinks, REPORT_TYPES.END_SESSION);
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'generateEndSessionReports', 
        'Error generating end-session reports. Please try again or contact support.');
    } else {
      Logger.log(`Error generating end-session reports: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to generate end-session reports: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Shows a UI for selecting classes
 * 
 * @param title - The title for the selection screen
 * @returns Array of selected class IDs
 */
function showClassSelectionUI(title) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      SpreadsheetApp.getUi().alert(
        'No Classes Found',
        'No classes were found in the system. Please add classes first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return [];
    }
    
    const classData = classesSheet.getDataRange().getValues();
    
    if (classData.length <= 1) {
      SpreadsheetApp.getUi().alert(
        'No Classes Found',
        'No classes were found in the system. Please add classes first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return [];
    }
    
    // Create a temporary UI for selection
    let tempSheet = ss.getSheetByName('TempClassSelection');
    if (!tempSheet) {
      tempSheet = ss.insertSheet('TempClassSelection');
    } else {
      tempSheet.clear();
    }
    
    // Add instructions and header
    tempSheet.getRange('A1:E1').merge()
      .setValue(title)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
      
    tempSheet.getRange('A2:E3').merge()
      .setValue('Check the boxes next to the classes you want to generate reports for. ' +
               'Then, click "Continue".')
      .setWrap(true);
    
    // Add class list with checkboxes
    tempSheet.getRange('A5').setValue('Select');
    tempSheet.getRange('B5').setValue('Class Name');
    tempSheet.getRange('C5').setValue('Day/Time');
    tempSheet.getRange('D5').setValue('Instructor');
    tempSheet.getRange('E5').setValue('Students');
    
    tempSheet.getRange('A5:E5')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add classes with checkboxes
    const classOptions = [];
    for (let i = 1; i < classData.length; i++) {
      const classId = classData[i][0];
      const className = classData[i][1];
      const day = classData[i][4];
      const time = classData[i][5];
      const instructor = classData[i][3];
      
      if (!classId || !className) continue;
      
      tempSheet.getRange(5 + classOptions.length, 1).insertCheckboxes();
      tempSheet.getRange(5 + classOptions.length, 2).setValue(className);
      tempSheet.getRange(5 + classOptions.length, 3).setValue(`${day} ${time}`);
      tempSheet.getRange(5 + classOptions.length, 4).setValue(instructor);
      
      // Count students in this class
      const studentCount = countStudentsInClass(classId);
      tempSheet.getRange(5 + classOptions.length, 5).setValue(studentCount);
      
      classOptions.push({
        classId: classId,
        row: 5 + classOptions.length
      });
    }
    
    // Add continue button
    tempSheet.getRange('C' + (8 + classOptions.length)).setValue('CONTINUE')
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Format sheet
    tempSheet.setColumnWidth(1, 80);
    tempSheet.setColumnWidth(2, 200);
    tempSheet.setColumnWidth(3, 150);
    tempSheet.setColumnWidth(4, 150);
    tempSheet.setColumnWidth(5, 100);
    
    // Activate sheet
    tempSheet.activate();
    
    // Wait for user to click continue
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Class Selection',
      'Select the classes you want to generate reports for, then click OK to continue.',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (result !== ui.Button.OK) {
      return [];
    }
    
    // Get selected classes
    const selectedClasses = [];
    for (const option of classOptions) {
      const isSelected = tempSheet.getRange(option.row, 1).getValue();
      if (isSelected) {
        selectedClasses.push(option.classId);
      }
    }
    
    return selectedClasses;
  } catch (error) {
    Logger.log(`Error showing class selection UI: ${error.message}`);
    return [];
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
 * Gets the URL of the report template for the specified report type
 * 
 * @param reportType - The type of report (mid-session or end-session)
 * @returns The template URL or null if not found
 */
function getReportTemplateUrl(reportType) {
  try {
    const config = AdministrativeModule.getSystemConfiguration();
    
    if (!config.reportTemplateUrl) {
      return null;
    }
    
    // Extract folder ID from URL
    let folderId = config.reportTemplateUrl;
    if (folderId.includes('/')) {
      const urlPattern = /[-\w]{25,}/;
      const match = folderId.match(urlPattern);
      if (match && match[0]) {
        folderId = match[0];
      }
    }
    
    // Look for template files in the folder
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFiles();
    
    // Expected file names
    const midSessionNames = [
      'Mid-Session Progress Report Template',
      'Mid Session Template',
      'Mid-Session Template',
      'Progress Report Template'
    ];
    
    const endSessionNames = [
      'End-Session Assessment Report Template',
      'End Session Template',
      'End-Session Template',
      'Assessment Report Template'
    ];
    
    // Search for matching template
    const targetNames = reportType === REPORT_TYPES.MID_SESSION ? midSessionNames : endSessionNames;
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      
      for (const name of targetNames) {
        if (fileName.toLowerCase().includes(name.toLowerCase())) {
          return file.getUrl();
        }
      }
    }
    
    return null;
  } catch (error) {
    Logger.log(`Error getting report template URL: ${error.message}`);
    return null;
  }
}

/**
 * Creates a folder for reports in Google Drive
 * 
 * @param folderName - The name of the folder to create
 * @returns The created folder
 */
function createReportFolder(folderName) {
  try {
    // Check if folder already exists
    const existingFolders = DriveApp.getFoldersByName(folderName);
    
    if (existingFolders.hasNext()) {
      return existingFolders.next();
    }
    
    // Create a new folder
    return DriveApp.createFolder(folderName);
  } catch (error) {
    Logger.log(`Error creating report folder: ${error.message}`);
    throw error;
  }
}

/**
 * Creates a subfolder within a parent folder
 * 
 * @param parentFolder - The parent folder
 * @param subfolderName - The name of the subfolder
 * @returns The created subfolder
 */
function createSubfolder(parentFolder, subfolderName) {
  try {
    // Check if folder already exists
    const existingFolders = parentFolder.getFoldersByName(subfolderName);
    
    if (existingFolders.hasNext()) {
      return existingFolders.next();
    }
    
    // Create a new folder
    return parentFolder.createFolder(subfolderName);
  } catch (error) {
    Logger.log(`Error creating subfolder: ${error.message}`);
    throw error;
  }
}

/**
 * Gets details for a class
 * 
 * @param classId - The class ID
 * @returns Object with class details
 */
function getClassDetails(classId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      return null;
    }
    
    const classData = classesSheet.getDataRange().getValues();
    
    // Skip header row, find matching class
    for (let i = 1; i < classData.length; i++) {
      if (classData[i][0] === classId) {
        return {
          classId: classId,
          className: classData[i][1],
          level: classData[i][2],
          instructor: classData[i][3],
          day: classData[i][4],
          time: classData[i][5],
          startDate: classData[i][6],
          endDate: classData[i][7],
          location: classData[i][8]
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
 * Gets students in a class with their assessment data
 * 
 * @param classId - The class ID
 * @returns Array of student objects with assessment data
 */
function getStudentsWithAssessments(classId) {
  try {
    const students = [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName('Roster');
    
    if (!rosterSheet) {
      return students;
    }
    
    const rosterData = rosterSheet.getDataRange().getValues();
    
    // Get tracker sheet for assessment data
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    if (!trackerSheet) {
      return students;
    }
    
    const trackerData = trackerSheet.getDataRange().getValues();
    
    // Find section in tracker that corresponds to this class
    let studentSection = -1;
    let skillsHeaderRow = -1;
    
    for (let i = 0; i < trackerData.length; i++) {
      // Look for the class info section
      if (trackerData[i][0] === 'Class ID:' && trackerData[i][1] === classId) {
        // Find the student data section (should be a few rows down)
        for (let j = i + 1; j < i + 20 && j < trackerData.length; j++) {
          if (trackerData[j][0] === 'Student ID' || trackerData[j][0] === 'Student') {
            studentSection = j + 1; // Start of student data
            skillsHeaderRow = j;
            break;
          }
        }
        break;
      }
    }
    
    if (studentSection === -1) {
      return students;
    }
    
    // Get skills from the header row
    const skills = [];
    for (let i = 2; i < trackerData[skillsHeaderRow].length; i++) {
      if (trackerData[skillsHeaderRow][i]) {
        skills.push({
          name: trackerData[skillsHeaderRow][i],
          index: i
        });
      }
    }
    
    // Extract students with their assessment data
    for (let i = studentSection; i < trackerData.length; i++) {
      const studentId = trackerData[i][0];
      const studentName = trackerData[i][1];
      
      if (!studentId || !studentName) {
        break; // End of student section
      }
      
      // Find student details in roster
      let studentEmail = '';
      let firstName = '';
      let lastName = '';
      let age = '';
      let level = '';
      let notes = '';
      
      for (let j = 1; j < rosterData.length; j++) {
        if (rosterData[j][0] === studentId) {
          studentEmail = rosterData[j][7]; // Assuming column 7 is parent email
          firstName = rosterData[j][2];    // Assuming column 2 is first name
          lastName = rosterData[j][3];     // Assuming column 3 is last name
          age = rosterData[j][4];          // Assuming column 4 is age
          level = rosterData[j][5];        // Assuming column 5 is level
          notes = rosterData[j][6];        // Assuming column 6 is notes
          break;
        }
      }
      
      // Get assessment data
      const assessments = [];
      for (const skill of skills) {
        assessments.push({
          skill: skill.name,
          assessment: trackerData[i][skill.index] || '',
          index: skill.index
        });
      }
      
      // Add student to list
      students.push({
        studentId: studentId,
        studentName: studentName,
        firstName: firstName || studentName.split(' ')[0],
        lastName: lastName || studentName.split(' ').slice(1).join(' '),
        email: studentEmail,
        age: age,
        level: level,
        notes: notes,
        assessments: assessments
      });
    }
    
    return students;
  } catch (error) {
    Logger.log(`Error getting students with assessments: ${error.message}`);
    return [];
  }
}

/**
 * Replaces placeholders in a document with student and class data
 * 
 * @param body - The document body
 * @param student - The student object
 * @param classDetails - The class details object
 * @param sessionName - The session name
 * @param nextLevel - Optional next level recommendation
 */
function replacePlaceholders(body, student, classDetails, sessionName, nextLevel = '') {
  try {
    // Basic replacements
    body.replaceText(PLACEHOLDERS.STUDENT_NAME, `${student.firstName} ${student.lastName}`);
    body.replaceText(PLACEHOLDERS.STUDENT_ID, student.studentId);
    body.replaceText(PLACEHOLDERS.CLASS_NAME, classDetails.className);
    body.replaceText(PLACEHOLDERS.LEVEL, student.level || classDetails.level);
    body.replaceText(PLACEHOLDERS.INSTRUCTOR_NAME, classDetails.instructor);
    body.replaceText(PLACEHOLDERS.SESSION_NAME, sessionName);
    body.replaceText(PLACEHOLDERS.REPORT_DATE, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMMM d, yyyy'));
    body.replaceText(PLACEHOLDERS.PARENT_NAME, 'Parent/Guardian');
    
    // Replace next level if provided
    if (nextLevel) {
      body.replaceText(PLACEHOLDERS.NEXT_LEVEL, nextLevel);
    }
    
    // Create skills table
    let skillsTable = '';
    if (student.assessments && student.assessments.length > 0) {
      skillsTable = 'Skill | Assessment\n-----|-------------\n';
      
      for (const assessment of student.assessments) {
        skillsTable += `${assessment.skill} | ${assessment.assessment}\n`;
      }
    }
    
    body.replaceText(PLACEHOLDERS.SKILLS_TABLE, skillsTable);
    
    // Replace comments with notes
    body.replaceText(PLACEHOLDERS.COMMENTS, student.notes || 'No additional comments.');
    
    // Next steps depends on assessment data
    let nextSteps = '';
    if (student.assessments && student.assessments.length > 0) {
      const incomplete = student.assessments.filter(a => 
        a.assessment !== 'P' && 
        a.assessment !== 'Proficient' && 
        a.assessment !== 'Pass' &&
        a.assessment !== '');
        
      if (incomplete.length > 0) {
        nextSteps = 'Focus areas for continued development:\n\n';
        for (const skill of incomplete) {
          nextSteps += `- ${skill.skill}\n`;
        }
      } else {
        nextSteps = 'Your child has completed all skills for this level! Congratulations!';
      }
    }
    
    body.replaceText(PLACEHOLDERS.NEXT_STEPS, nextSteps);
  } catch (error) {
    Logger.log(`Error replacing placeholders: ${error.message}`);
    throw error;
  }
}

/**
 * Calculates the next level recommendation for a student
 * 
 * @param student - The student object with assessment data
 * @returns The recommended next level
 */
function calculateNextLevel(student) {
  try {
    // Get current level
    const currentLevel = student.level;
    
    if (!currentLevel) {
      return 'Next appropriate level';
    }
    
    // Calculate proficiency percentage
    let totalSkills = 0;
    let proficientSkills = 0;
    
    for (const assessment of student.assessments) {
      if (assessment.skill && assessment.assessment) {
        totalSkills++;
        
        if (assessment.assessment === 'P' || 
            assessment.assessment === 'Proficient' || 
            assessment.assessment === 'Pass') {
          proficientSkills++;
        }
      }
    }
    
    // If 80% or more skills are proficient, recommend next level
    if (totalSkills > 0 && proficientSkills / totalSkills >= 0.8) {
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
      
      return levelProgression[currentLevel] || 'Next appropriate level';
    } else {
      // Recommend repeating current level
      return currentLevel + ' (repeat for further skill development)';
    }
  } catch (error) {
    Logger.log(`Error calculating next level: ${error.message}`);
    return 'Next appropriate level';
  }
}

/**
 * Generates an assessment summary text
 * 
 * @param student - The student object with assessment data
 * @returns Assessment summary text
 */
function generateAssessmentSummary(student) {
  try {
    // Calculate percentages
    let totalSkills = 0;
    let proficientSkills = 0;
    let developingSkills = 0;
    let notAttemptedSkills = 0;
    
    for (const assessment of student.assessments) {
      if (assessment.skill) {
        totalSkills++;
        
        if (assessment.assessment === 'P' || 
            assessment.assessment === 'Proficient' || 
            assessment.assessment === 'Pass') {
          proficientSkills++;
        } else if (assessment.assessment === 'D' || 
                  assessment.assessment === 'Developing' || 
                  assessment.assessment === 'In Progress') {
          developingSkills++;
        } else if (!assessment.assessment) {
          notAttemptedSkills++;
        }
      }
    }
    
    const proficientPercent = totalSkills > 0 ? Math.round(proficientSkills / totalSkills * 100) : 0;
    
    // Generate summary text
    let summary = `${student.firstName} has completed ${proficientSkills} out of ${totalSkills} skills (${proficientPercent}%) `;
    
    if (proficientPercent >= 80) {
      summary += 'and has demonstrated strong proficiency in this level. ';
      summary += 'Based on this assessment, your child is ready to advance to the next level.';
    } else if (proficientPercent >= 50) {
      summary += 'and has made good progress toward mastering this level. ';
      summary += 'Several skills still need further development before advancing to the next level.';
    } else {
      summary += 'and will benefit from additional practice to develop the core skills of this level. ';
      summary += 'We recommend continuing at this level for the next session.';
    }
    
    return summary;
  } catch (error) {
    Logger.log(`Error generating assessment summary: ${error.message}`);
    return 'Assessment summary not available.';
  }
}

/**
 * Generates class recommendations for the next session
 * 
 * @param student - The student object
 * @param nextLevel - The recommended next level
 * @returns Recommendations text
 */
function generateRecommendations(student, nextLevel) {
  try {
    if (!nextLevel || nextLevel.includes('repeat')) {
      return `We recommend registering for ${student.level} for the next session to continue developing the necessary skills.`;
    } else {
      return `Based on this assessment, we recommend registering for ${nextLevel} for the next session. This will provide appropriate challenges to continue ${student.firstName}'s swimming development.`;
    }
  } catch (error) {
    Logger.log(`Error generating recommendations: ${error.message}`);
    return 'Please check with the aquatics department for recommendations on your child\'s next swimming level.';
  }
}

/**
 * Creates a summary sheet with links to all generated reports
 * 
 * @param reportLinks - Array of report link objects
 * @param errors - Array of error messages
 * @param sheetName - The name for the summary sheet
 * @param folderUrl - The URL of the reports folder
 */
function createReportSummarySheet(reportLinks, errors, sheetName, folderUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let summarySheet = ss.getSheetByName(sheetName);
    
    if (!summarySheet) {
      summarySheet = ss.insertSheet(sheetName);
    } else {
      summarySheet.clear();
    }
    
    // Add header
    summarySheet.getRange('A1:E1').merge()
      .setValue('Report Generation Summary')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Add folder link
    summarySheet.getRange('A2:E2').merge()
      .setValue(`Reports Folder: ${folderUrl}`)
      .setWrap(true);
    
    // Add generation timestamp
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'MMMM d, yyyy h:mm a');
    summarySheet.getRange('A3:E3').merge()
      .setValue(`Generated on: ${timestamp}`)
      .setWrap(true);
    
    // Add report list headers
    summarySheet.getRange('A5').setValue('Student Name');
    summarySheet.getRange('B5').setValue('Class');
    summarySheet.getRange('C5').setValue('Report Link');
    summarySheet.getRange('D5').setValue('Next Level');
    summarySheet.getRange('E5').setValue('Email Sent');
    
    summarySheet.getRange('A5:E5')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add report links
    for (let i = 0; i < reportLinks.length; i++) {
      const link = reportLinks[i];
      summarySheet.getRange(6 + i, 1).setValue(link.student);
      summarySheet.getRange(6 + i, 2).setValue(link.class);
      
      // Add hyperlink
      summarySheet.getRange(6 + i, 3).setValue('View Report')
        .setFontColor('blue')
        .setTextStyle(SpreadsheetApp.newTextStyle().setUnderline(true).build());
      
      // Set the hyperlink
      summarySheet.getRange(6 + i, 3).setFormula(`=HYPERLINK("${link.url}", "View Report")`);
      
      // Add next level if available
      if (link.nextLevel) {
        summarySheet.getRange(6 + i, 4).setValue(link.nextLevel);
      }
      
      // Email status (will be updated later if emails are sent)
      summarySheet.getRange(6 + i, 5).setValue('No');
    }
    
    // Add errors section if there are any
    if (errors.length > 0) {
      const errorRow = 7 + reportLinks.length;
      
      summarySheet.getRange(errorRow, 1, 1, 5).merge()
        .setValue('Errors and Warnings')
        .setFontWeight('bold')
        .setBackground('#FFC107')
        .setFontColor('black');
      
      for (let i = 0; i < errors.length; i++) {
        summarySheet.getRange(errorRow + 1 + i, 1, 1, 5).merge()
          .setValue(errors[i])
          .setWrap(true);
      }
    }
    
    // Format sheet
    summarySheet.setColumnWidth(1, 200);
    summarySheet.setColumnWidth(2, 200);
    summarySheet.setColumnWidth(3, 100);
    summarySheet.setColumnWidth(4, 200);
    summarySheet.setColumnWidth(5, 100);
    
    // Activate sheet
    summarySheet.activate();
    
    return summarySheet;
  } catch (error) {
    Logger.log(`Error creating summary sheet: ${error.message}`);
    // Don't throw, this is a non-critical operation
  }
}

/**
 * Emails reports to parents
 * 
 * @param reportLinks - Array of report link objects
 * @param reportType - The type of report
 * @returns Success status
 */
function emailReports(reportLinks, reportType) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Emailing reports to parents', 'INFO', 'emailReports');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Check if all reports have email addresses
    const noEmailReports = reportLinks.filter(r => !r.email);
    
    if (noEmailReports.length > 0) {
      // Ask user what to do about missing emails
      const missingResult = ui.alert(
        'Missing Email Addresses',
        `${noEmailReports.length} reports do not have parent email addresses. Would you like to continue sending emails to the others?`,
        ui.ButtonSet.YES_NO
      );
      
      if (missingResult !== ui.Button.YES) {
        return false;
      }
    }
    
    // Get email template
    const emailSubject = reportType === REPORT_TYPES.MID_SESSION ? 
      'Mid-Session Progress Report' : 'End-Session Assessment Report';
    
    let emailBody = 'Dear Parent/Guardian,\n\n' +
      'We are pleased to provide you with your child\'s ' + 
      (reportType === REPORT_TYPES.MID_SESSION ? 'mid-session progress report' : 'end-session assessment report') + 
      ' for their swim lessons.\n\n' +
      'You can access the report using the link below:\n\n' +
      '{{REPORT_LINK}}\n\n' +
      'Please let us know if you have any questions about your child\'s progress.\n\n' +
      'Thank you,\nYMCA Aquatics Team';
    
    // Ask if user wants to customize the email
    const customizeResult = ui.alert(
      'Customize Email',
      'Would you like to customize the email message?',
      ui.ButtonSet.YES_NO
    );
    
    if (customizeResult === ui.Button.YES) {
      const emailResult = ui.prompt(
        'Customize Email',
        'Edit the email message (use {{REPORT_LINK}} as a placeholder for the report link):',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (emailResult.getSelectedButton() === ui.Button.OK && emailResult.getResponseText()) {
        emailBody = emailResult.getResponseText();
      }
    }
    
    // Send emails
    let successCount = 0;
    let failCount = 0;
    const errors = [];
    
    for (const report of reportLinks) {
      // Skip reports without emails
      if (!report.email) {
        continue;
      }
      
      try {
        // Replace the link placeholder
        const personalizedBody = emailBody.replace('{{REPORT_LINK}}', report.url);
        
        // Send the email
        GmailApp.sendEmail(
          report.email,
          `${emailSubject} - ${report.student}`,
          personalizedBody,
          {
            name: 'YMCA Swim Lessons'
          }
        );
        
        successCount++;
        
        // Update the summary sheet
        updateReportSummary(report.student, true);
      } catch (emailError) {
        failCount++;
        errors.push(`Error sending to ${report.student}: ${emailError.message}`);
        
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Error sending report email to ${report.email}: ${emailError.message}`, 'ERROR', 'emailReports');
        }
      }
    }
    
    // Log communication
    if (CommunicationModule && typeof CommunicationModule.createCommunicationLog === 'function') {
      // Ensure log sheet exists
      CommunicationModule.createCommunicationLog();
      
      // Add to log
      if (successCount > 0) {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const logSheet = ss.getSheetByName('CommunicationLog');
        
        if (logSheet) {
          const now = new Date();
          const date = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');
          const time = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm:ss');
          
          const status = failCount > 0 ? 'Partial' : 'Sent';
          const notes = failCount > 0 ? 
            `${successCount} sent, ${failCount} failed. Reports sent via email.` : 
            `${successCount} sent successfully. Reports sent via email.`;
          
          // Add log entry
          logSheet.insertRowAfter(1);
          logSheet.getRange(2, 1, 1, 8).setValues([[
            date,
            time,
            reportType,
            `${successCount + failCount} recipients`,
            emailSubject,
            Session.getEffectiveUser().getEmail(),
            status,
            notes
          ]]);
        }
      }
    }
    
    // Show results
    if (failCount > 0) {
      ui.alert(
        'Emails Sent with Errors',
        `Successfully sent ${successCount} emails with ${failCount} failures.\n\nSee error details in the log.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Emails Sent',
        `Successfully sent ${successCount} emails.`,
        ui.ButtonSet.OK
      );
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'emailReports', 
        'Error sending report emails. Please try again or contact support.');
    } else {
      Logger.log(`Error sending report emails: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to send report emails: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Updates the report summary sheet with email status
 * 
 * @param studentName - The student name
 * @param emailSent - Whether the email was sent
 */
function updateReportSummary(studentName, emailSent) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check both summary sheet names
    const sheetNames = [
      `${REPORT_TYPES.MID_SESSION} - Summary`,
      `${REPORT_TYPES.END_SESSION} - Summary`
    ];
    
    for (const name of sheetNames) {
      const sheet = ss.getSheetByName(name);
      
      if (sheet) {
        // Find the student row
        const dataRange = sheet.getDataRange();
        const data = dataRange.getValues();
        
        for (let i = 5; i < data.length; i++) {
          if (data[i][0] === studentName) {
            // Update the email sent column
            sheet.getRange(i + 1, 5).setValue(emailSent ? 'Yes' : 'No');
            break;
          }
        }
      }
    }
  } catch (error) {
    Logger.log(`Error updating report summary: ${error.message}`);
    // Non-critical, so don't throw
  }
}

// Global variable export
const ReportingModule = {
  generateMidSessionReports,
  generateEndSessionReports
};