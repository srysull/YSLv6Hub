/**
 * YSL Hub Dynamic Instructor Sheet Module
 * 
 * This module implements a new approach to instructor sheets with a single dynamic
 * sheet that updates based on class selection. It pulls student data from rosters
 * and skills data from the Swimmer Records Workbook.
 * 
 * @author Claude Code
 * @version 1.0
 * @date 2025-05-05
 */

// Configuration constants
const DYNAMIC_INSTRUCTOR_CONFIG = {
  SHEET_NAME: 'Instructor Sheet',
  ROSTER_SHEET_NAME: 'Daxko', // Changed from 'Roster' to 'Daxko'
  HEADERS: {
    CLASS_SELECTOR_LABEL: 'Select Class:',
    FIRST_NAME: 'First Name',
    LAST_NAME: 'Last Name',
    ATTENDANCE_PREFIX: 'Day ',
    ATTENDANCE_COUNT: 8,
  },
  CELL_STYLES: {
    HEADER_COLOR: '#4285F4',
    HEADER_TEXT_COLOR: '#FFFFFF',
    SECTION_COLOR: '#E0E0E0',
    ATTENDANCE_COLOR: '#F0F8FF',
    STAGE_SKILLS_COLOR: '#F0FFF0',
    SAW_SKILLS_COLOR: '#FFF0F0'
  },
  DAXKO_COLUMNS: {
    FIRST_NAME: 2, // Column C (0-indexed)
    LAST_NAME: 3,  // Column D
    PROGRAM: 22,   // Column W
    SESSION: 23    // Column X
  }
};

/**
 * Creates or resets the dynamic instructor sheet
 * @return {Sheet} The created or updated sheet
 */
function createDynamicInstructorSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    
    // Create the sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    } else {
      // Clear existing content but keep the sheet
      sheet.clear();
    }
    
    // Set up the basic structure
    setupSheetStructure(sheet);
    
    // Create class selector dropdown
    createClassSelector(sheet);
    
    // Add onEdit trigger for the sheet
    ensureTriggerExists();
    
    // Show confirmation
    SpreadsheetApp.getUi().alert(
      'Dynamic Instructor Sheet Created',
      'The dynamic instructor sheet has been created. Select a class from the dropdown to load student data.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Set active sheet to instructor sheet
    sheet.activate();
    
    return sheet;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createDynamicInstructorSheet', 
        'Error creating dynamic instructor sheet. Please try again or contact support.');
    } else {
      Logger.log(`Error creating dynamic instructor sheet: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create dynamic instructor sheet: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return null;
  }
}

/**
 * Sets up the basic structure of the instructor sheet
 * @param {Sheet} sheet - The sheet to set up
 */
function setupSheetStructure(sheet) {
  // Set column widths
  sheet.setColumnWidth(1, 150); // First Name
  sheet.setColumnWidth(2, 150); // Last Name
  
  // Set up basic headers and format
  sheet.getRange('A1:B1').merge()
    .setValue(DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.CLASS_SELECTOR_LABEL)
    .setFontWeight('bold');
  
  sheet.getRange('C1:D1').merge(); // Space for class selector dropdown
  
  // Format header row for student info (will be filled when class is selected)
  sheet.getRange('A3').setValue(DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.FIRST_NAME)
    .setFontWeight('bold')
    .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
    .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
  
  sheet.getRange('B3').setValue(DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.LAST_NAME)
    .setFontWeight('bold')
    .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
    .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
  
  // Freeze the header rows
  sheet.setFrozenRows(3);
}

/**
 * Creates the class selector dropdown
 * @param {Sheet} sheet - The instructor sheet
 */
function createClassSelector(sheet) {
  try {
    // Get available classes
    const classes = getAvailableClasses();
    
    if (classes.length === 0) {
      Logger.log('No classes found for the selector');
      return;
    }
    
    // Create a dropdown with all available classes
    const classRange = sheet.getRange('C1:D1');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(classes, true)
      .build();
    
    classRange.setDataValidation(rule);
    
    // Add a note to explain how to use the dropdown
    classRange.setNote('Select a class to load student data and skills. The sheet will update automatically.');
  } catch (error) {
    Logger.log(`Error creating class selector: ${error.message}`);
    throw error;
  }
}

/**
 * Gets all available classes from the Classes sheet
 * @return {Array} Array of class names (Program Day Time)
 */
function getAvailableClasses() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    if (!classesSheet) {
      throw new Error('Classes sheet not found');
    }
    
    // Get class data
    const classData = classesSheet.getDataRange().getValues();
    
    // Skip header row
    const classNames = [];
    for (let i = 1; i < classData.length; i++) {
      // Check if row has valid data
      if (classData[i][1] && classData[i][2] && classData[i][3]) {
        const className = `${classData[i][1]} ${classData[i][2]} ${classData[i][3]}`;
        classNames.push(className);
      }
    }
    
    return classNames;
  } catch (error) {
    Logger.log(`Error getting available classes: ${error.message}`);
    throw error;
  }
}

/**
 * Ensures an onEdit trigger exists for the spreadsheet
 */
function ensureTriggerExists() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let hasClassSelectorTrigger = false;
    
    // Check if trigger already exists
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onEditDynamicInstructorSheet') {
        hasClassSelectorTrigger = true;
        break;
      }
    }
    
    // Create trigger if it doesn't exist
    if (!hasClassSelectorTrigger) {
      ScriptApp.newTrigger('onEditDynamicInstructorSheet')
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onEdit()
        .create();
    }
  } catch (error) {
    Logger.log(`Error ensuring trigger exists: ${error.message}`);
    // Continue even if trigger creation fails
  }
}

/**
 * Handles edits to the instructor sheet, particularly class selection
 * @param {Object} e - The onEdit event object
 */
function onEditDynamicInstructorSheet(e) {
  try {
    // Check if edit was in the instructor sheet
    if (!e || !e.range || e.range.getSheet().getName() !== DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME) {
      return;
    }
    
    // Check if the edit was in the class selector (C1:D1)
    if (e.range.getRow() === 1 && (e.range.getColumn() === 3 || e.range.getColumn() === 4)) {
      const selectedClass = e.value;
      
      if (!selectedClass) {
        return; // No class selected
      }
      
      // Populate the sheet with the selected class's data
      populateSheetWithClassData(e.range.getSheet(), selectedClass);
    }
  } catch (error) {
    Logger.log(`Error in onEditDynamicInstructorSheet: ${error.message}`);
    // Don't throw errors in trigger functions
  }
}

/**
 * Populates the instructor sheet with data for the selected class
 * @param {Sheet} sheet - The instructor sheet
 * @param {string} selectedClass - The selected class (Program Day Time)
 */
function populateSheetWithClassData(sheet, selectedClass) {
  try {
    // Clear existing student data
    clearStudentData(sheet);
    
    // Get class details from the selected class
    const classDetails = parseClassDetails(selectedClass);
    
    // Get student roster for this class
    const students = getStudentsForClass(classDetails);
    
    if (students.length === 0) {
      sheet.getRange('A4').setValue('No students found for this class');
      return;
    }
    
    // Add attendance columns
    setupAttendanceColumns(sheet);
    
    // Get skills from swimmer records
    const skills = getSkillsFromSwimmerRecords();
    
    // Add skills columns
    setupSkillsColumns(sheet, skills);
    
    // Populate student data
    populateStudentData(sheet, students, skills);
    
    // Add class header
    sheet.getRange('A2:Z2').merge()
      .setValue(`Class: ${selectedClass}`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR)
      .setHorizontalAlignment('center');
  } catch (error) {
    Logger.log(`Error populating sheet with class data: ${error.message}`);
    sheet.getRange('A4').setValue(`Error loading class data: ${error.message}`);
    throw error;
  }
}

/**
 * Clears existing student data from the sheet
 * @param {Sheet} sheet - The instructor sheet
 */
function clearStudentData(sheet) {
  // Get the data range (excluding headers)
  const lastRow = Math.max(sheet.getLastRow(), 100);
  const lastCol = Math.max(sheet.getLastColumn(), 100);
  
  if (lastRow > 3) {
    sheet.getRange(4, 1, lastRow - 3, lastCol).clear();
  }
}

/**
 * Parses class details from the selected class string
 * @param {string} selectedClass - The selected class (Program Day Time)
 * @return {Object} The class details object
 */
function parseClassDetails(selectedClass) {
  try {
    // Parse out program, day, and time
    const parts = selectedClass.split(' ');
    
    // The program may have multiple words, so we need to find where the day starts
    const days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
    let dayIndex = -1;
    
    for (let i = 0; i < parts.length; i++) {
      if (days.includes(parts[i])) {
        dayIndex = i;
        break;
      }
    }
    
    // If day not found, assume format is different
    if (dayIndex === -1) {
      return {
        fullName: selectedClass,
        program: selectedClass,
        day: '',
        time: ''
      };
    }
    
    const program = parts.slice(0, dayIndex).join(' ');
    const day = parts[dayIndex];
    const time = parts.slice(dayIndex + 1).join(' ');
    
    return {
      fullName: selectedClass,
      program: program,
      day: day,
      time: time
    };
  } catch (error) {
    Logger.log(`Error parsing class details: ${error.message}`);
    return {
      fullName: selectedClass,
      program: selectedClass,
      day: '',
      time: ''
    };
  }
}

/**
 * Gets students for the specified class from the roster
 * @param {Object} classDetails - The class details
 * @return {Array} Array of student objects
 */
function getStudentsForClass(classDetails) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rosterSheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME);
    
    if (!rosterSheet) {
      throw new Error(`${DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME} sheet not found. Please make sure the Daxko sheet exists.`);
    }
    
    // Get all roster data
    const rosterData = rosterSheet.getDataRange().getValues();
    
    // Filter students for this class
    const students = [];
    for (let i = 1; i < rosterData.length; i++) {
      // Match by program and session in columns W and X (22 and 23 in 0-indexed)
      const rowProgram = rosterData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM];
      
      // Check if this row matches the class
      if (rowProgram && rowProgram.includes(classDetails.program)) {
        // Make sure we have valid first and last names
        const firstName = rosterData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.FIRST_NAME];
        const lastName = rosterData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.LAST_NAME];
        
        if (firstName && lastName) {
          students.push({
            firstName: firstName,
            lastName: lastName,
            skills: {} // Will be populated later with skills from the Swimmer Records
          });
        }
      }
    }
    
    if (students.length === 0) {
      Logger.log(`No students found matching program: ${classDetails.program}`);
    } else {
      Logger.log(`Found ${students.length} students for program: ${classDetails.program}`);
    }
    
    return students;
  } catch (error) {
    Logger.log(`Error getting students for class: ${error.message}`);
    throw error;
  }
}

/**
 * Sets up attendance columns in the sheet
 * @param {Sheet} sheet - The instructor sheet
 */
function setupAttendanceColumns(sheet) {
  // Add attendance column headers
  for (let i = 1; i <= DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT; i++) {
    const col = i + 2; // Start at column C (after first & last name)
    sheet.getRange(3, col).setValue(`${DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_PREFIX}${i}`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.ATTENDANCE_COLOR);
    
    // Set column width
    sheet.setColumnWidth(col, 80);
  }
}

/**
 * Gets all skill headers from the Swimmer Records Workbook
 * @return {Object} Object with skill headers categorized by type
 */
function getSkillsFromSwimmerRecords() {
  try {
    // Get Swimmer Records URL from properties
    const swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL);
    if (!swimmerRecordsUrl) {
      throw new Error('Swimmer Records URL not found in system configuration');
    }
    
    // Extract spreadsheet ID from URL
    const ssId = GlobalFunctions.extractIdFromUrl(swimmerRecordsUrl);
    if (!ssId) {
      throw new Error('Invalid Swimmer Records URL');
    }
    
    // Open the Swimmer Records Workbook
    const swimmerSS = SpreadsheetApp.openById(ssId);
    const swimmerSheet = swimmerSS.getSheets()[0]; // Assuming first sheet contains the records
    
    // Get the header row
    const headerRow = swimmerSheet.getRange(1, 1, 1, swimmerSheet.getLastColumn()).getValues()[0];
    
    // Categorize skills
    const skills = {
      stage: [], // For stage skills (prefixed with S1, S2, etc.)
      saw: []    // For SAW skills (prefixed with SAW)
    };
    
    // Start from column 3 (after first and last name)
    for (let i = 2; i < headerRow.length; i++) {
      const header = headerRow[i];
      if (!header) continue;
      
      // Check skill type by prefix
      const headerStr = header.toString();
      if (headerStr.startsWith('S') && !headerStr.startsWith('SAW')) {
        skills.stage.push({
          index: i,
          header: headerStr
        });
      } else if (headerStr.startsWith('SAW')) {
        skills.saw.push({
          index: i,
          header: headerStr
        });
      }
    }
    
    return skills;
  } catch (error) {
    Logger.log(`Error getting skills from Swimmer Records: ${error.message}`);
    throw error;
  }
}

/**
 * Sets up skills columns in the sheet
 * @param {Sheet} sheet - The instructor sheet
 * @param {Object} skills - The skills object
 */
function setupSkillsColumns(sheet, skills) {
  // Calculate starting column for skills (after attendance columns)
  const startCol = 3 + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT;
  
  // Add stage skills headers
  for (let i = 0; i < skills.stage.length; i++) {
    const col = startCol + i;
    sheet.getRange(3, col).setValue(skills.stage[i].header)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.STAGE_SKILLS_COLOR);
    
    // Set column width
    sheet.setColumnWidth(col, 100);
  }
  
  // Add SAW skills headers
  const sawStartCol = startCol + skills.stage.length;
  for (let i = 0; i < skills.saw.length; i++) {
    const col = sawStartCol + i;
    sheet.getRange(3, col).setValue(skills.saw[i].header)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SAW_SKILLS_COLOR);
    
    // Set column width
    sheet.setColumnWidth(col, 100);
  }
}

/**
 * Populates student data and existing skills
 * @param {Sheet} sheet - The instructor sheet
 * @param {Array} students - Array of student objects
 * @param {Object} skills - Skills configuration
 */
function populateStudentData(sheet, students, skills) {
  try {
    // Get student skills from Swimmer Records
    const studentSkills = getStudentSkillsFromSwimmerRecords(students);
    
    // Add students to the sheet
    for (let i = 0; i < students.length; i++) {
      const rowIndex = i + 4; // Start at row 4 (after headers)
      
      // Add student name
      sheet.getRange(rowIndex, 1).setValue(students[i].firstName);
      sheet.getRange(rowIndex, 2).setValue(students[i].lastName);
      
      // Leave attendance columns empty
      
      // Add existing skills if available
      const skillsStartCol = 3 + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT;
      
      // Find student in skillsData
      const student = studentSkills.find(s => 
        s.firstName === students[i].firstName && 
        s.lastName === students[i].lastName);
      
      if (student) {
        // Add stage skills
        for (let j = 0; j < skills.stage.length; j++) {
          const skillHeader = skills.stage[j].header;
          if (student.skills[skillHeader]) {
            sheet.getRange(rowIndex, skillsStartCol + j).setValue(student.skills[skillHeader]);
          }
        }
        
        // Add SAW skills
        const sawStartCol = skillsStartCol + skills.stage.length;
        for (let j = 0; j < skills.saw.length; j++) {
          const skillHeader = skills.saw[j].header;
          if (student.skills[skillHeader]) {
            sheet.getRange(rowIndex, sawStartCol + j).setValue(student.skills[skillHeader]);
          }
        }
      }
    }
  } catch (error) {
    Logger.log(`Error populating student data: ${error.message}`);
    throw error;
  }
}

/**
 * Gets existing student skills from the Swimmer Records Workbook
 * @param {Array} students - Array of student objects
 * @return {Array} Array of students with their skills
 */
function getStudentSkillsFromSwimmerRecords(students) {
  try {
    // Get Swimmer Records URL from properties
    const swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL);
    if (!swimmerRecordsUrl) {
      throw new Error('Swimmer Records URL not found in system configuration');
    }
    
    // Extract spreadsheet ID from URL
    const ssId = GlobalFunctions.extractIdFromUrl(swimmerRecordsUrl);
    if (!ssId) {
      throw new Error('Invalid Swimmer Records URL');
    }
    
    // Open the Swimmer Records Workbook
    const swimmerSS = SpreadsheetApp.openById(ssId);
    const swimmerSheet = swimmerSS.getSheets()[0]; // Assuming first sheet contains the records
    
    // Get all data
    const recordsData = swimmerSheet.getDataRange().getValues();
    if (recordsData.length <= 1) {
      return []; // No data or only headers
    }
    
    // Get headers
    const headers = recordsData[0];
    
    // Find skills for each student
    const studentSkills = [];
    
    for (let i = 1; i < recordsData.length; i++) {
      const firstName = recordsData[i][0];
      const lastName = recordsData[i][1];
      
      // Check if this student is in our roster
      const matchingStudent = students.find(s => 
        s.firstName === firstName && 
        s.lastName === lastName);
      
      if (matchingStudent) {
        const skills = {};
        
        // Collect all skills for this student
        for (let j = 2; j < headers.length; j++) {
          const header = headers[j];
          if (!header) continue;
          
          const value = recordsData[i][j];
          if (value) {
            skills[header] = value;
          }
        }
        
        studentSkills.push({
          firstName: firstName,
          lastName: lastName,
          skills: skills
        });
      }
    }
    
    return studentSkills;
  } catch (error) {
    Logger.log(`Error getting student skills from Swimmer Records: ${error.message}`);
    return []; // Return empty array on error
  }
}

/**
 * Pushes skills updates back to the Swimmer Records Workbook
 */
function pushSkillsToSwimmerRecords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const instructorSheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    
    if (!instructorSheet) {
      throw new Error('Instructor sheet not found');
    }
    
    // Check if a class is selected
    const selectedClass = instructorSheet.getRange('C1:D1').getValue();
    if (!selectedClass) {
      throw new Error('No class selected. Please select a class first.');
    }
    
    // Get all data from the instructor sheet
    const instructorData = instructorSheet.getDataRange().getValues();
    
    // Get headers (row 3)
    const headers = instructorData[2];
    
    // Get Swimmer Records URL from properties
    const swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL);
    if (!swimmerRecordsUrl) {
      throw new Error('Swimmer Records URL not found in system configuration');
    }
    
    // Extract spreadsheet ID from URL
    const ssId = GlobalFunctions.extractIdFromUrl(swimmerRecordsUrl);
    if (!ssId) {
      throw new Error('Invalid Swimmer Records URL');
    }
    
    // Open the Swimmer Records Workbook
    const swimmerSS = SpreadsheetApp.openById(ssId);
    const swimmerSheet = swimmerSS.getSheets()[0]; // Assuming first sheet contains the records
    
    // Get all Swimmer Records data
    const recordsData = swimmerSheet.getDataRange().getValues();
    const recordsHeaders = recordsData[0];
    
    // Track changes
    let updatedCount = 0;
    let newSkillsCount = 0;
    
    // Process each student in the instructor sheet (starting at row 4)
    for (let i = 3; i < instructorData.length; i++) {
      const firstName = instructorData[i][0];
      const lastName = instructorData[i][1];
      
      if (!firstName || !lastName) continue;
      
      // Find this student in the Swimmer Records
      let studentFound = false;
      let studentRow = -1;
      
      for (let j = 1; j < recordsData.length; j++) {
        if (recordsData[j][0] === firstName && recordsData[j][1] === lastName) {
          studentFound = true;
          studentRow = j;
          break;
        }
      }
      
      if (!studentFound) {
        // If student not in Records, log but continue
        Logger.log(`Student not found in Records: ${firstName} ${lastName}`);
        continue;
      }
      
      // Look at all skill columns (after attendance)
      const skillsStartCol = 3 + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT;
      
      for (let j = skillsStartCol; j < headers.length; j++) {
        const skillHeader = headers[j];
        if (!skillHeader) continue;
        
        // Find this skill in the Records sheet
        const recordsCol = recordsHeaders.indexOf(skillHeader);
        if (recordsCol === -1) continue;
        
        // Check if there's a value to update
        const value = instructorData[i][j];
        if (!value) continue;
        
        // Allow only valid skill values ('X' or '/')
        if (value === 'X' || value === '/') {
          // Check if this is a new value or update
          if (recordsData[studentRow][recordsCol] !== value) {
            // Update the cell in Swimmer Records
            swimmerSheet.getRange(studentRow + 1, recordsCol + 1).setValue(value);
            
            if (recordsData[studentRow][recordsCol]) {
              updatedCount++;
            } else {
              newSkillsCount++;
            }
          }
        }
      }
    }
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      'Skills Updated',
      `Skills have been pushed to Swimmer Records Workbook.\n\n` +
      `${updatedCount} skills updated\n` +
      `${newSkillsCount} new skills added`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'pushSkillsToSwimmerRecords', 
        'Error pushing skills to Swimmer Records. Please try again or contact support.');
    } else {
      Logger.log(`Error pushing skills to Swimmer Records: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to push skills to Swimmer Records: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

// Make functions available to other modules
const DynamicInstructorSheet = {
  createDynamicInstructorSheet: createDynamicInstructorSheet,
  pushSkillsToSwimmerRecords: pushSkillsToSwimmerRecords,
  onEditDynamicInstructorSheet: onEditDynamicInstructorSheet
};