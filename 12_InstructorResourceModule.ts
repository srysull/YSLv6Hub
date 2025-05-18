/**
 * YSL Hub v2 Instructor Resource Module
 * 
 * This module provides resources and tools for swim instructors, including
 * generating instructor-specific sheets, providing access to teaching
 * materials, and facilitating class management.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

/**
 * Generates instructor-specific sheets for all classes
 * 
 * @returns Success status
 */
function generateInstructorSheets() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Generating instructor sheets', 'INFO', 'generateInstructorSheets');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Check if classes exist
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      ui.alert(
        'Missing Data',
        'The Classes sheet is missing. Please initialize the system first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    const classData = classesSheet.getDataRange().getValues();
    
    if (classData.length <= 1) {
      ui.alert(
        'No Classes',
        'There are no classes defined in the system. Please add classes first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Show selection UI
    const instructors = getUniqueInstructors(classData);
    
    if (instructors.length === 0) {
      ui.alert(
        'No Instructors',
        'No instructors were found in the class data. Please make sure classes have assigned instructors.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Create a temporary sheet for selection
    let selectSheet = ss.getSheetByName('InstructorSelection');
    if (selectSheet) {
      selectSheet.clear();
    } else {
      selectSheet = ss.insertSheet('InstructorSelection');
    }
    
    // Set up selection sheet header
    selectSheet.getRange('A1:D1').merge()
      .setValue('Generate Instructor Sheets')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    selectSheet.getRange('A2:D4').merge()
      .setValue('Select instructors to generate sheets for. Check the boxes next to the instructors you want to include.')
      .setWrap(true);
    
    // Add instructor list with checkboxes
    selectSheet.getRange('A6').setValue('Select');
    selectSheet.getRange('B6').setValue('Instructor Name');
    selectSheet.getRange('C6').setValue('Number of Classes');
    selectSheet.getRange('D6').setValue('Status');
    
    selectSheet.getRange('A6:D6')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add instructors with counts
    for (let i = 0; i < instructors.length; i++) {
      const instructor = instructors[i];
      selectSheet.getRange(7 + i, 1).insertCheckboxes();
      selectSheet.getRange(7 + i, 2).setValue(instructor);
      
      // Count classes taught by this instructor
      let classCount = 0;
      for (let j = 1; j < classData.length; j++) {
        if (classData[j][3] === instructor) { // Assuming column 3 is Instructor
          classCount++;
        }
      }
      
      selectSheet.getRange(7 + i, 3).setValue(classCount);
      selectSheet.getRange(7 + i, 4).setValue('Not Generated');
    }
    
    // Add continue button
    selectSheet.getRange(7 + instructors.length + 2, 2).setValue('CONTINUE')
      .setBackground('#4285F4')
      .setFontColor('white')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Format sheet
    selectSheet.setColumnWidth(1, 80);
    selectSheet.setColumnWidth(2, 200);
    selectSheet.setColumnWidth(3, 120);
    selectSheet.setColumnWidth(4, 120);
    
    // Activate sheet
    selectSheet.activate();
    
    // Wait for user to click continue
    const result = ui.alert(
      'Generate Instructor Sheets',
      'Select the instructors you want to generate sheets for, then click OK to continue.',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (result !== ui.Button.OK) {
      return false;
    }
    
    // Get selected instructors
    const selectedInstructors = [];
    for (let i = 0; i < instructors.length; i++) {
      const isSelected = selectSheet.getRange(7 + i, 1).getValue();
      if (isSelected) {
        selectedInstructors.push(instructors[i]);
      }
    }
    
    if (selectedInstructors.length === 0) {
      ui.alert(
        'No Selection',
        'Please select at least one instructor.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Process each selected instructor
    let successCount = 0;
    let failCount = 0;
    
    for (let i = 0; i < selectedInstructors.length; i++) {
      const instructor = selectedInstructors[i];
      
      // Update status
      selectSheet.getRange(7 + instructors.indexOf(instructor), 4).setValue('Generating...');
      
      // Generate instructor sheet
      const success = createInstructorSheet(instructor, classData);
      
      if (success) {
        successCount++;
        selectSheet.getRange(7 + instructors.indexOf(instructor), 4).setValue('Generated')
          .setBackground('#4CAF50')
          .setFontColor('white');
      } else {
        failCount++;
        selectSheet.getRange(7 + instructors.indexOf(instructor), 4).setValue('Failed')
          .setBackground('#F44336')
          .setFontColor('white');
      }
    }
    
    // Show results
    if (failCount > 0) {
      ui.alert(
        'Generation Complete',
        `Successfully generated ${successCount} instructor sheets with ${failCount} failures.`,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Generation Complete',
        `Successfully generated ${successCount} instructor sheets.`,
        ui.ButtonSet.OK
      );
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'generateInstructorSheets', 
        'Error generating instructor sheets. Please try again or contact support.');
    } else {
      Logger.log(`Error generating instructor sheets: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to generate instructor sheets: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Gets a list of unique instructors from class data
 * 
 * @param classData - The class data array
 * @returns Array of unique instructor names
 */
function getUniqueInstructors(classData) {
  const instructors = new Set();
  
  // Skip header row
  for (let i = 1; i < classData.length; i++) {
    const instructor = classData[i][3]; // Assuming column 3 is Instructor
    if (instructor) {
      instructors.add(instructor);
    }
  }
  
  return Array.from(instructors);
}

/**
 * Creates a sheet for a specific instructor with their class information
 * 
 * @param instructorName - The name of the instructor
 * @param classData - The class data array
 * @returns Success status
 */
function createInstructorSheet(instructorName, classData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = `Instructor - ${instructorName}`;
    
    // Check if sheet already exists
    let instructorSheet = ss.getSheetByName(sheetName);
    if (instructorSheet) {
      instructorSheet.clear();
    } else {
      instructorSheet = ss.insertSheet(sheetName);
    }
    
    // Get system info
    const config = AdministrativeModule.getSystemConfiguration();
    const sessionName = config.sessionName || 'Current Session';
    
    // Set up header
    instructorSheet.getRange('A1:G1').merge()
      .setValue(`Instructor Sheet: ${instructorName}`)
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    instructorSheet.getRange('A2:G2').merge()
      .setValue(`Session: ${sessionName}`)
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Add classes section
    instructorSheet.getRange('A4:G4').merge()
      .setValue('Your Classes')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add class headers
    instructorSheet.getRange('A5').setValue('Class Name');
    instructorSheet.getRange('B5').setValue('Level');
    instructorSheet.getRange('C5').setValue('Day');
    instructorSheet.getRange('D5').setValue('Time');
    instructorSheet.getRange('E5').setValue('Location');
    instructorSheet.getRange('F5').setValue('Students');
    instructorSheet.getRange('G5').setValue('Notes');
    
    instructorSheet.getRange('A5:G5')
      .setFontWeight('bold')
      .setBackground('#E1E1E1');
    
    // Filter classes for this instructor
    const instructorClasses = [];
    for (let i = 1; i < classData.length; i++) {
      if (classData[i][3] === instructorName) { // Assuming column 3 is Instructor
        instructorClasses.push(classData[i]);
      }
    }
    
    // Add class data
    for (let i = 0; i < instructorClasses.length; i++) {
      const rowData = [
        instructorClasses[i][1], // Class Name
        instructorClasses[i][2], // Level
        instructorClasses[i][4], // Day
        instructorClasses[i][5], // Time
        instructorClasses[i][8], // Location
        instructorClasses[i][9], // Students
        instructorClasses[i][12]  // Notes
      ];
      
      instructorSheet.getRange(6 + i, 1, 1, 7).setValues([rowData]);
      
      // Add alternating row background
      if (i % 2 === 1) {
        instructorSheet.getRange(6 + i, 1, 1, 7).setBackground('#f9f9f9');
      }
    }
    
    // Get student data for this instructor's classes
    if (instructorClasses.length > 0) {
      addStudentRoster(instructorSheet, instructorClasses, 8 + instructorClasses.length);
    }
    
    // Add teaching resources section
    const startRow = 10 + instructorClasses.length + getStudentCount(instructorClasses);
    addTeachingResources(instructorSheet, startRow);
    
    // Format sheet
    instructorSheet.setColumnWidth(1, 200); // Class Name
    instructorSheet.setColumnWidth(2, 100); // Level
    instructorSheet.setColumnWidth(3, 80);  // Day
    instructorSheet.setColumnWidth(4, 80);  // Time
    instructorSheet.setColumnWidth(5, 120); // Location
    instructorSheet.setColumnWidth(6, 80);  // Students
    instructorSheet.setColumnWidth(7, 200); // Notes
    
    return true;
  } catch (error) {
    Logger.log(`Error creating instructor sheet for ${instructorName}: ${error.message}`);
    return false;
  }
}

/**
 * Adds student roster information to the instructor sheet
 * 
 * @param sheet - The instructor sheet
 * @param instructorClasses - The instructor's classes
 * @param startRow - The starting row for the roster section
 */
function addStudentRoster(sheet, instructorClasses, startRow) {
  try {
    // Add roster section header
    sheet.getRange(startRow, 1, 1, 7).merge()
      .setValue('Student Roster')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add roster headers
    sheet.getRange(startRow + 1, 1).setValue('Class');
    sheet.getRange(startRow + 1, 2).setValue('Student Name');
    sheet.getRange(startRow + 1, 3).setValue('Age');
    sheet.getRange(startRow + 1, 4).setValue('Level');
    sheet.getRange(startRow + 1, 5).setValue('Email');
    sheet.getRange(startRow + 1, 6).setValue('Phone');
    sheet.getRange(startRow + 1, 7).setValue('Notes');
    
    sheet.getRange(startRow + 1, 1, 1, 7)
      .setFontWeight('bold')
      .setBackground('#E1E1E1');
    
    // Get roster data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName('Roster');
    
    if (!rosterSheet) {
      return;
    }
    
    const rosterData = rosterSheet.getDataRange().getValues();
    
    if (rosterData.length <= 1) {
      return;
    }
    
    // Extract class IDs for this instructor
    const classIds = instructorClasses.map(classRow => classRow[0]); // Assuming column 0 is Class ID
    
    // Filter students for these classes
    let currentRow = startRow + 2;
    
    for (let i = 1; i < rosterData.length; i++) {
      const classId = rosterData[i][1]; // Assuming column 1 is Class ID in roster
      
      if (classIds.includes(classId)) {
        // Find class name
        let className = '';
        for (const classRow of instructorClasses) {
          if (classRow[0] === classId) {
            className = classRow[1]; // Assuming column 1 is Class Name
            break;
          }
        }
        
        // Add student data
        const rowData = [
          className,
          `${rosterData[i][2]} ${rosterData[i][3]}`, // First + Last name
          rosterData[i][4], // Age
          rosterData[i][5], // Level
          rosterData[i][7], // Email
          rosterData[i][8], // Phone
          rosterData[i][6]  // Notes
        ];
        
        sheet.getRange(currentRow, 1, 1, 7).setValues([rowData]);
        
        // Add alternating row background
        if ((currentRow - startRow - 2) % 2 === 1) {
          sheet.getRange(currentRow, 1, 1, 7).setBackground('#f9f9f9');
        }
        
        currentRow++;
      }
    }
  } catch (error) {
    Logger.log(`Error adding student roster: ${error.message}`);
  }
}

/**
 * Gets the count of students in the instructor's classes
 * 
 * @param instructorClasses - The instructor's classes
 * @returns The number of students
 */
function getStudentCount(instructorClasses) {
  try {
    // Extract class IDs for this instructor
    const classIds = instructorClasses.map(classRow => classRow[0]); // Assuming column 0 is Class ID
    
    // Get roster data
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName('Roster');
    
    if (!rosterSheet) {
      return 0;
    }
    
    const rosterData = rosterSheet.getDataRange().getValues();
    
    if (rosterData.length <= 1) {
      return 0;
    }
    
    // Count students in these classes
    let count = 0;
    
    for (let i = 1; i < rosterData.length; i++) {
      const classId = rosterData[i][1]; // Assuming column 1 is Class ID in roster
      
      if (classIds.includes(classId)) {
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
 * Adds teaching resources to the instructor sheet
 * 
 * @param sheet - The instructor sheet
 * @param startRow - The starting row for the resources section
 */
function addTeachingResources(sheet, startRow) {
  try {
    // Add resources section header
    sheet.getRange(startRow, 1, 1, 7).merge()
      .setValue('Teaching Resources')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Get config for resource URLs
    const config = AdministrativeModule.getSystemConfiguration();
    const parentHandbookUrl = config.parentHandbookUrl || '';
    
    // Add resources section content
    const resources = [
      ['Parent Handbook', parentHandbookUrl, 'Information for parents about the swim program'],
      ['Lesson Plans', '', 'Access lesson plans for each level'],
      ['Assessment Guide', '', 'Guidelines for assessing student skills'],
      ['Safety Protocols', '', 'Important safety information for instructors'],
      ['Instructor FAQs', '', 'Common questions and answers for instructors']
    ];
    
    for (let i = 0; i < resources.length; i++) {
      sheet.getRange(startRow + 1 + i, 1).setValue(resources[i][0]);
      
      if (resources[i][1]) {
        // Add hyperlink
        sheet.getRange(startRow + 1 + i, 2).setValue('VIEW')
          .setFontColor('blue')
          .setTextStyle(SpreadsheetApp.newTextStyle().setUnderline(true).build());
        
        // Set hyperlink formula
        sheet.getRange(startRow + 1 + i, 2).setFormula(`=HYPERLINK("${resources[i][1]}", "VIEW")`);
      } else {
        sheet.getRange(startRow + 1 + i, 2).setValue('N/A');
      }
      
      sheet.getRange(startRow + 1 + i, 3, 1, 5).merge()
        .setValue(resources[i][2]);
    }
    
    // Add other important information
    sheet.getRange(startRow + 7, 1, 1, 7).merge()
      .setValue('Important Information for Instructors')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    const infoText = [
      'Please ensure you arrive 15 minutes before your class starts.',
      'If you need to request a substitute, please contact the Aquatics Coordinator at least 24 hours in advance.',
      'All assessments should be completed by the second-to-last class of the session.',
      'Remember to complete progress reports for each student.',
      'If you have questions or need assistance, please contact the Aquatics Department.'
    ];
    
    for (let i = 0; i < infoText.length; i++) {
      sheet.getRange(startRow + 8 + i, 1, 1, 7).merge()
        .setValue(infoText[i]);
    }
  } catch (error) {
    Logger.log(`Error adding teaching resources: ${error.message}`);
  }
}

// Global variable export
const InstructorResourceModule = {
  generateInstructorSheets
};