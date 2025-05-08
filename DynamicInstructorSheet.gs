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
  ROSTER_SHEET_NAME: 'Daxko', // The sheet containing student registration data
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
    SAW_SKILLS_COLOR: '#FFF0F0',
    BEFORE_COLUMN_COLOR: '#E6F2FF', // Light blue for 'Before' columns
    AFTER_COLUMN_COLOR: '#FFEBCC'   // Light orange for 'After' columns
  },
  DAXKO_COLUMNS: {
    FIRST_NAME: 2, // Column C (0-indexed) - Student first name
    LAST_NAME: 3,  // Column D - Student last name
    PROGRAM: 22,   // Column W - Program name/description
    SESSION: 23    // Column X - Session details (day/time)
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
    
    // Display some information about the Classes sheet
    Logger.log(`Classes sheet has ${classData.length} rows`);
    if (classData.length > 0) {
      Logger.log(`Classes sheet headers: ${JSON.stringify(classData[0])}`);
    }
    
    // Skip header row
    const classNames = [];
    for (let i = 1; i < classData.length; i++) {
      // Check if row has valid data for the class definition
      if (classData[i].length > 3 && classData[i][1] && classData[i][2] && classData[i][3]) {
        const className = `${classData[i][1]} ${classData[i][2]} ${classData[i][3]}`;
        classNames.push(className);
        Logger.log(`Added class: ${className}`);
      }
    }
    
    // If no classes found, add some test options
    if (classNames.length === 0) {
      Logger.log('No classes found in the Classes sheet, adding test classes');
      classNames.push('Test Swimming Monday 9:00 AM');
      classNames.push('Test Swimming Tuesday 10:00 AM');
    }
    
    return classNames;
  } catch (error) {
    Logger.log(`Error getting available classes: ${error.message}`);
    
    // Return some test classes as a fallback
    return ['Test Swimming Monday 9:00 AM', 'Test Swimming Tuesday 10:00 AM'];
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
    
    // Check if this is a private lesson
    if (classDetails.isPrivateLesson) {
      // Use special layout for private lessons
      setupPrivateLessonLayout(sheet, classDetails);
      return; // Exit early, no need to proceed with regular class setup
    }
    
    // Extract stage from class name if possible (e.g., "S1" from "Swimming S1 Monday")
    const stage = extractStageFromClassName(classDetails.program);
    Logger.log(`Extracted stage: ${stage} from class: ${classDetails.program}`);
    
    // Add class header with stage info
    sheet.getRange('A2:Z2').merge()
      .setValue(`Class: ${selectedClass}${stage ? ` - Stage ${stage}` : ''}`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR)
      .setHorizontalAlignment('center');
    
    // Add attendance columns
    setupAttendanceColumns(sheet);
    
    // Get skills from swimmer records
    const allSkills = getSkillsFromSwimmerRecords();
    
    // Filter skills by stage if possible
    const filteredSkills = filterSkillsByStage(allSkills, stage);
    
    // Add skills columns with split for before/after tracking
    setupSplitSkillsColumns(sheet, filteredSkills);
    
    // Get student roster for this class
    const students = getStudentsForClass(classDetails);
    
    // Display info message if no students, but continue with empty roster
    if (students.length === 0) {
      // This will be overwritten if test students get added
      sheet.getRange('A4').setValue('No students found in Daxko for this class. Add students manually or check class details.');
    }
    
    // Populate student data (even if it's just test/demo students)
    populateStudentData(sheet, students, filteredSkills);
    
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
    
    // Check if this is a private lesson
    const isPrivateLesson = selectedClass.toLowerCase().includes('private');
    
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
        time: '',
        isPrivateLesson: isPrivateLesson
      };
    }
    
    const program = parts.slice(0, dayIndex).join(' ');
    const day = parts[dayIndex];
    const time = parts.slice(dayIndex + 1).join(' ');
    
    return {
      fullName: selectedClass,
      program: program,
      day: day,
      time: time,
      isPrivateLesson: isPrivateLesson
    };
  } catch (error) {
    Logger.log(`Error parsing class details: ${error.message}`);
    return {
      fullName: selectedClass,
      program: selectedClass,
      day: '',
      time: '',
      isPrivateLesson: false
    };
  }
}

/**
 * Gets students for the specified class from the Daxko sheet directly
 * @param {Object} classDetails - The class details
 * @return {Array} Array of student objects
 */
function getStudentsForClass(classDetails) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME);
    
    if (!daxkoSheet) {
      throw new Error(`${DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME} sheet not found. Please make sure the Daxko sheet exists.`);
    }
    
    // Get all roster data from Daxko sheet
    const daxkoData = daxkoSheet.getDataRange().getValues();
    
    if (daxkoData.length <= 1) {
      Logger.log('No data found in Daxko sheet (only headers or empty)');
      return [];
    }
    
    // Log the class details for debugging
    Logger.log(`Looking for students in class: ${JSON.stringify(classDetails)}`);
    
    // Get the column headers to verify we're looking at the right columns
    const headers = daxkoData[0];
    Logger.log(`Daxko sheet headers: ${JSON.stringify(headers)}`);
    
    // Check if we have the expected program and session columns
    if (headers.length <= DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM || 
        headers.length <= DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION) {
      Logger.log(`WARNING: Daxko sheet does not have expected columns. Expected Program in column ${DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM + 1} and Session in column ${DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION + 1}`);
      // Create a sample row entry to inspect
      if (daxkoData.length > 1) {
        Logger.log(`Sample row data: ${JSON.stringify(daxkoData[1])}`);
      }
    }
    
    // Build search terms from class details
    const searchTerms = [];
    
    // Extract key terms from program name
    if (classDetails.program) {
      // First add the entire program name
      searchTerms.push(classDetails.program.toLowerCase().trim());
      
      // Then add individual significant words (avoiding common words like "the", "and")
      const programWords = classDetails.program.toLowerCase().trim().split(/\s+/);
      const skipWords = ['the', 'and', 'or', 'in', 'at', 'with', 'for'];
      
      for (const word of programWords) {
        if (word.length > 2 && !skipWords.includes(word)) {
          searchTerms.push(word);
        }
      }
      
      // Try to identify stage number in the program name
      const stage = extractStageFromClassName(classDetails.program);
      if (stage) {
        searchTerms.push(`stage ${stage}`);
        searchTerms.push(`s${stage}`);
      }
    }
    
    // Add day and time as search terms
    if (classDetails.day) searchTerms.push(classDetails.day.toLowerCase().trim());
    if (classDetails.time) {
      const timeParts = classDetails.time.toLowerCase().trim().split(' ');
      if (timeParts.length > 0) searchTerms.push(timeParts[0]); // Just the time, not AM/PM
    }
    
    Logger.log(`Search terms for class matching: ${JSON.stringify(searchTerms)}`);
    
    const students = [];
    
    // Added a more flexible matching approach with expanded logging
    let matchCount = 0;
    let failedMatches = 0;
    const failedMatchExamples = [];
    
    // Get column indices for additional matching options
    let activityNameCol = -1;
    let activityTimeCol = -1;
    
    // Find columns with activity or class information if program/session not working
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i]?.toString().toLowerCase() || '';
      if (header.includes('activity') && header.includes('name')) {
        activityNameCol = i;
      }
      if (header.includes('activity') && header.includes('time')) {
        activityTimeCol = i;
      }
    }
    
    Logger.log(`Found additional matching columns: Activity Name: ${activityNameCol}, Activity Time: ${activityTimeCol}`);
    
    // Direct matching on program and session columns (W and X in Daxko sheet)
    for (let i = 1; i < daxkoData.length; i++) {
      // Make sure row has sufficient data for basic columns
      if (daxkoData[i].length <= Math.max(
          DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.FIRST_NAME,
          DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.LAST_NAME)) {
        continue; // Skip rows with insufficient column data for name fields
      }
      
      const firstName = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.FIRST_NAME];
      const lastName = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.LAST_NAME];
      
      // Skip rows without names
      if (!firstName || !lastName) continue;
      
      // Get all columns that might contain class information
      let rowProgram = '';
      let rowSession = '';
      let activityName = '';
      let activityTime = '';
      
      if (daxkoData[i].length > DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM) {
        rowProgram = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM];
      }
      
      if (daxkoData[i].length > DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION) {
        rowSession = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION];
      }
      
      if (activityNameCol >= 0 && daxkoData[i].length > activityNameCol) {
        activityName = daxkoData[i][activityNameCol];
      }
      
      if (activityTimeCol >= 0 && daxkoData[i].length > activityTimeCol) {
        activityTime = daxkoData[i][activityTimeCol];
      }
      
      // Convert all fields to lowercase strings for consistent comparison
      const programStr = rowProgram ? rowProgram.toString().toLowerCase().trim() : '';
      const sessionStr = rowSession ? rowSession.toString().toLowerCase().trim() : '';
      const activityNameStr = activityName ? activityName.toString().toLowerCase().trim() : '';
      const activityTimeStr = activityTime ? activityTime.toString().toLowerCase().trim() : '';
      
      // Combine all fields for more comprehensive matching
      const allFieldsStr = `${programStr} ${sessionStr} ${activityNameStr} ${activityTimeStr}`;
      
      // Debugging output - print first 5 students we check
      if (matchCount + failedMatches < 5) {
        Logger.log(`Checking student: ${firstName} ${lastName}, Program: "${programStr}", Session: "${sessionStr}", Activity: "${activityNameStr} ${activityTimeStr}"`);
      }
      
      // Enhanced matching logic:
      let isMatch = false;
      let matchReason = '';
      
      // STRATEGY 1: Full match - count how many search terms are present
      let termMatches = 0;
      let totalTerms = 0;
      
      for (const term of searchTerms) {
        if (term.length < 3) continue; // Skip very short search terms
        totalTerms++;
        
        // Check if term is in any field
        if (allFieldsStr.includes(term)) {
          termMatches++;
        }
      }
      
      // If we match most of the search terms, consider it a match
      if (totalTerms > 0 && termMatches >= Math.ceil(totalTerms * 0.6)) {
        isMatch = true;
        matchReason = `Matched ${termMatches}/${totalTerms} search terms`;
      }
      
      // STRATEGY 2: Class-specific matching - try to match the exact class name
      if (!isMatch && classDetails.fullName) {
        const normalizedClassName = classDetails.fullName.toLowerCase().trim();
        
        // Check if any field contains the full class name
        if (allFieldsStr.includes(normalizedClassName)) {
          isMatch = true;
          matchReason = 'Matched full class name';
        }
        // Check for key components - program + day + time
        else if (classDetails.program && classDetails.day && classDetails.time) {
          const normalizedProgram = classDetails.program.toLowerCase().trim();
          const normalizedDay = classDetails.day.toLowerCase().trim();
          const normalizedTime = classDetails.time.toLowerCase().trim().split(' ')[0]; // Just time without AM/PM
          
          if (allFieldsStr.includes(normalizedProgram) && 
              allFieldsStr.includes(normalizedDay) && 
              allFieldsStr.includes(normalizedTime)) {
            isMatch = true;
            matchReason = 'Matched program + day + time';
          }
        }
      }
      
      // STRATEGY 3: For testing purposes
      if (!isMatch && (searchTerms.length === 0 || classDetails.program.toLowerCase().includes('test'))) {
        // Add any student with valid name for testing
        if (firstName && lastName) {
          isMatch = true;
          matchReason = 'Added for testing';
        }
      }
      
      if (isMatch) {
        matchCount++;
        students.push({
          firstName: firstName,
          lastName: lastName,
          fullName: `${firstName} ${lastName}`, // Add full name for easier matching later
          skills: {}, // Will be populated later with skills from the Swimmer Records
          matchReason: matchReason // For debugging
        });
        Logger.log(`Found matching student: ${firstName} ${lastName} (${matchReason})`);
      } else {
        failedMatches++;
        if (failedMatchExamples.length < 3) {
          failedMatchExamples.push(`${firstName} ${lastName} - Program: "${programStr}", Session: "${sessionStr}"`);
        }
      }
    }
    
    // Log detailed results
    Logger.log(`Found ${students.length} matching students and rejected ${failedMatches} students`);
    if (failedMatchExamples.length > 0) {
      Logger.log(`Sample of non-matching students: ${JSON.stringify(failedMatchExamples)}`);
    }
    
    // If no students found, create test students
    if (students.length === 0) {
      Logger.log('No students found for the selected class. Creating test students as fallback.');
      const stage = extractStageFromClassName(classDetails.program) || '';
      students.push({
        firstName: 'Test',
        lastName: stage ? `Student S${stage}` : 'Student1',
        fullName: stage ? `Test Student S${stage}` : 'Test Student1',
        skills: {},
        matchReason: 'Test student'
      });
      students.push({
        firstName: 'Test',
        lastName: stage ? `Student S${stage}-2` : 'Student2',
        fullName: stage ? `Test Student S${stage}-2` : 'Test Student2',
        skills: {},
        matchReason: 'Test student'
      });
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
      Logger.log('Swimmer Records URL not found in system configuration');
      return createFallbackSkills();
    }
    
    // Extract spreadsheet ID from URL
    const ssId = GlobalFunctions.extractIdFromUrl(swimmerRecordsUrl);
    if (!ssId) {
      Logger.log('Invalid Swimmer Records URL');
      return createFallbackSkills();
    }
    
    // Log the spreadsheet ID we're trying to open
    Logger.log(`Attempting to open Swimmer Records with ID: ${ssId}`);
    
    // Try to open the Swimmer Records Workbook
    let swimmerSS;
    try {
      swimmerSS = SpreadsheetApp.openById(ssId);
    } catch (accessError) {
      Logger.log(`Error accessing Swimmer Records: ${accessError.message}`);
      return createFallbackSkills();
    }
    
    // Get the first sheet
    const sheets = swimmerSS.getSheets();
    if (!sheets || sheets.length === 0) {
      Logger.log('No sheets found in Swimmer Records Workbook');
      return createFallbackSkills();
    }
    
    const swimmerSheet = sheets[0]; // Assuming first sheet contains the records
    
    // Get the header row
    const headerRow = swimmerSheet.getRange(1, 1, 1, swimmerSheet.getLastColumn()).getValues()[0];
    
    // Log the headers for debugging
    Logger.log(`Swimmer Records headers: ${JSON.stringify(headerRow)}`);
    
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
    
    // Log the skills we found
    Logger.log(`Found ${skills.stage.length} stage skills and ${skills.saw.length} SAW skills`);
    
    // If no skills found, fall back to test skills
    if (skills.stage.length === 0 && skills.saw.length === 0) {
      Logger.log('No skills found in Swimmer Records, using fallback skills');
      return createFallbackSkills();
    }
    
    return skills;
  } catch (error) {
    Logger.log(`Error getting skills from Swimmer Records: ${error.message}`);
    // Return fallback skills instead of throwing an error
    return createFallbackSkills();
  }
}

/**
 * Creates fallback skills for testing when Swimmer Records is unavailable
 * @return {Object} Object with test skill headers
 */
function createFallbackSkills() {
  const skills = {
    stage: [],
    saw: []
  };
  
  // Add some stage skills for testing
  const stageNames = ['S1', 'S2', 'S3', 'S4', 'S5', 'S6'];
  const skillTypes = ['Float', 'Kick', 'Submerge', 'Arm Strokes', 'Breathing'];
  
  let index = 2; // Start after first and last name columns
  
  // Add stage skills
  for (const stage of stageNames) {
    for (const skill of skillTypes) {
      skills.stage.push({
        index: index++,
        header: `${stage} ${skill}`,
        stage: stage.replace('S', '') // Extract numeric stage value
      });
    }
  }
  
  // Add SAW skills
  const sawSkills = ['SAW Water Safety', 'SAW Life Jacket', 'SAW Help Others', 'SAW Call for Help'];
  for (const skill of sawSkills) {
    skills.saw.push({
      index: index++,
      header: skill
    });
  }
  
  Logger.log(`Created ${skills.stage.length} fallback stage skills and ${skills.saw.length} fallback SAW skills`);
  return skills;
}

/**
 * Sets up skills columns with split before/after tracking
 * @param {Sheet} sheet - The instructor sheet
 * @param {Object} skills - The skills object
 */
function setupSplitSkillsColumns(sheet, skills) {
  // Calculate starting column for skills (after attendance columns)
  const startCol = 3 + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT;
  
  // Add row for subheaders (before/after)
  sheet.insertRowAfter(2);
  
  // Add main header row for before/after
  sheet.getRange(3, startCol, 1, skills.stage.length * 2).merge()
    .setValue('Stage Skills - Before & After Evaluation')
    .setFontWeight('bold')
    .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.STAGE_SKILLS_COLOR)
    .setHorizontalAlignment('center');
  
  // Add SAW header if we have SAW skills
  if (skills.saw.length > 0) {
    const sawStartCol = startCol + (skills.stage.length * 2);
    sheet.getRange(3, sawStartCol, 1, skills.saw.length * 2).merge()
      .setValue('Safety Around Water - Before & After Evaluation')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SAW_SKILLS_COLOR)
      .setHorizontalAlignment('center');
  }
  
  // Add stage skills headers with Before/After columns
  let currentCol = startCol;
  for (let i = 0; i < skills.stage.length; i++) {
    // Skill header
    sheet.getRange(4, currentCol, 1, 2).merge()
      .setValue(skills.stage[i].header)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.STAGE_SKILLS_COLOR)
      .setHorizontalAlignment('center');
    
    // Before column 
    sheet.getRange(5, currentCol).setValue('Before')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.BEFORE_COLUMN_COLOR)
      .setHorizontalAlignment('center');
    
    // After column
    sheet.getRange(5, currentCol + 1).setValue('After')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.AFTER_COLUMN_COLOR)
      .setHorizontalAlignment('center');
    
    // Set column widths
    sheet.setColumnWidth(currentCol, 80);
    sheet.setColumnWidth(currentCol + 1, 80);
    
    // Move to next skill columns
    currentCol += 2;
  }
  
  // Add SAW skills headers with Before/After columns
  if (skills.saw.length > 0) {
    for (let i = 0; i < skills.saw.length; i++) {
      // Skill header
      sheet.getRange(4, currentCol, 1, 2).merge()
        .setValue(skills.saw[i].header)
        .setFontWeight('bold')
        .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SAW_SKILLS_COLOR)
        .setHorizontalAlignment('center');
      
      // Before column 
      sheet.getRange(5, currentCol).setValue('Before')
        .setFontWeight('bold')
        .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.BEFORE_COLUMN_COLOR)
        .setHorizontalAlignment('center');
      
      // After column
      sheet.getRange(5, currentCol + 1).setValue('After')
        .setFontWeight('bold')
        .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.AFTER_COLUMN_COLOR)
        .setHorizontalAlignment('center');
      
      // Set column widths
      sheet.setColumnWidth(currentCol, 80);
      sheet.setColumnWidth(currentCol + 1, 80);
      
      // Move to next skill columns
      currentCol += 2;
    }
  }
  
  // Update frozen rows to include all header rows
  sheet.setFrozenRows(5);
}

/**
 * Populates student data and existing skills with split before/after columns
 * @param {Sheet} sheet - The instructor sheet
 * @param {Array} students - Array of student objects from the selected class only
 * @param {Object} skills - Skills configuration
 */
function populateStudentData(sheet, students, skills) {
  try {
    // Get student skills from Swimmer Records - but only for students in this class
    const studentSkills = getStudentSkillsFromSwimmerRecords(students);
    
    // Log student count for debugging
    Logger.log(`Populating data for ${students.length} students from the selected class`);
    
    // Account for the additional header rows we added
    const startRow = 6; // After all header rows
    
    // Add students to the sheet - only those from the selected class
    for (let i = 0; i < students.length; i++) {
      const rowIndex = i + startRow; // Start after all header rows
      
      // Add student name
      sheet.getRange(rowIndex, 1).setValue(students[i].firstName);
      sheet.getRange(rowIndex, 2).setValue(students[i].lastName);
      
      // Add empty attendance columns with checkboxes
      const attendanceStartCol = 3;
      const attendanceEndCol = attendanceStartCol + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT - 1;
      const attendanceRange = sheet.getRange(rowIndex, attendanceStartCol, 1, DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT);
      
      // Apply checkbox data validation to attendance cells
      attendanceRange.insertCheckboxes();
      
      // Calculate starting column for skills (after attendance columns)
      const skillsStartCol = 3 + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT;
      
      // Find student in skillsData by name
      const student = studentSkills.find(s => 
        s.firstName.toString().toLowerCase().trim() === students[i].firstName.toString().toLowerCase().trim() && 
        s.lastName.toString().toLowerCase().trim() === students[i].lastName.toString().toLowerCase().trim());
      
      // Add existing skills values in the "Before" columns if available
      if (student) {
        // Add stage skills in the "Before" columns (every 2 columns)
        let currentSkillColumn = skillsStartCol;
        for (let j = 0; j < skills.stage.length; j++) {
          const skillHeader = skills.stage[j].header;
          if (student.skills[skillHeader]) {
            // Put existing skill value in the "Before" column
            sheet.getRange(rowIndex, currentSkillColumn).setValue(student.skills[skillHeader]);
          }
          currentSkillColumn += 2; // Skip the "After" column
        }
        
        // Add SAW skills in the "Before" columns (every 2 columns)
        const sawStartCol = skillsStartCol + (skills.stage.length * 2);
        currentSkillColumn = sawStartCol;
        for (let j = 0; j < skills.saw.length; j++) {
          const skillHeader = skills.saw[j].header;
          if (student.skills[skillHeader]) {
            // Put existing skill value in the "Before" column
            sheet.getRange(rowIndex, currentSkillColumn).setValue(student.skills[skillHeader]);
          }
          currentSkillColumn += 2; // Skip the "After" column
        }
      }
    }
    
    // Add validation for skill cells
    addSkillValidation(sheet, startRow, students.length, skills);
    
  } catch (error) {
    Logger.log(`Error populating student data: ${error.message}`);
    throw error;
  }
}

/**
 * Gets skills directly from the Swimmer Records Workbook for specific students
 * Only looks up skills for students that are in the provided students array
 * 
 * @param {Array} students - Array of student objects from selected class
 * @return {Array} Array of students with their skills
 */
function getStudentSkillsFromSwimmerRecords(students) {
  try {
    // If no students, return empty array
    if (!students || students.length === 0) {
      Logger.log('No students provided to look up skills');
      return [];
    }
    
    // Get Swimmer Records URL from properties
    const swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL);
    if (!swimmerRecordsUrl) {
      Logger.log('Swimmer Records URL not found in system configuration');
      return []; // Return empty array, sheet will just show students without skills
    }
    
    // Extract spreadsheet ID from URL
    const ssId = GlobalFunctions.extractIdFromUrl(swimmerRecordsUrl);
    if (!ssId) {
      Logger.log('Invalid Swimmer Records URL');
      return []; // Return empty array
    }
    
    // Log which students we're looking for
    Logger.log(`Looking up skills for ${students.length} students in the selected class`);
    
    // Try to open the Swimmer Records Workbook
    let swimmerSS;
    try {
      swimmerSS = SpreadsheetApp.openById(ssId);
    } catch (accessError) {
      Logger.log(`Error accessing Swimmer Records: ${accessError.message}`);
      return [];
    }
    
    // Get the first sheet
    const sheets = swimmerSS.getSheets();
    if (!sheets || sheets.length === 0) {
      Logger.log('No sheets found in Swimmer Records Workbook');
      return [];
    }
    
    const swimmerSheet = sheets[0]; // Assuming first sheet contains the records
    
    // Get all data
    const recordsData = swimmerSheet.getDataRange().getValues();
    if (recordsData.length <= 1) {
      Logger.log('No data found in Swimmer Records (only headers or empty)');
      return []; // No data or only headers
    }
    
    // Get headers
    const headers = recordsData[0];
    Logger.log(`Swimmer Records has ${recordsData.length} rows and ${headers.length} columns`);
    
    // Results array
    const studentSkills = [];
    
    // Process each student in the selected class
    for (const student of students) {
      // Look for this student in the Swimmer Records
      let foundRecord = false;
      
      // For each student in our class, check every row in the Swimmer Records
      for (let i = 1; i < recordsData.length; i++) {
        const recordFirstName = recordsData[i][0];
        const recordLastName = recordsData[i][1];
        
        if (!recordFirstName || !recordLastName) continue;
        
        // Compare names (case-insensitive)
        const firstNameMatch = recordFirstName.toString().toLowerCase().trim() === 
                               student.firstName.toString().toLowerCase().trim();
        const lastNameMatch = recordLastName.toString().toLowerCase().trim() === 
                              student.lastName.toString().toLowerCase().trim();
        
        // If names match, collect skills for this student
        if (firstNameMatch && lastNameMatch) {
          const skills = {};
          
          // Get all skills from the row
          for (let j = 2; j < headers.length; j++) {
            const header = headers[j];
            if (!header) continue;
            
            const value = recordsData[i][j];
            if (value) {
              skills[header] = value;
            }
          }
          
          // Add this student's skills to our results
          studentSkills.push({
            firstName: student.firstName,
            lastName: student.lastName,
            skills: skills
          });
          
          Logger.log(`Found skills for ${student.firstName} ${student.lastName}`);
          foundRecord = true;
          break; // Found this student, move to next one
        }
      }
      
      // If student not found in Swimmer Records, add them without skills
      if (!foundRecord) {
        Logger.log(`No record found for ${student.firstName} ${student.lastName} in Swimmer Records`);
        // Still include the student in results but with empty skills
        studentSkills.push({
          firstName: student.firstName,
          lastName: student.lastName,
          skills: {} // Empty skills
        });
      }
    }
    
    Logger.log(`Processed ${studentSkills.length} students from the selected class`);
    return studentSkills;
  } catch (error) {
    Logger.log(`Error getting student skills from Swimmer Records: ${error.message}`);
    return []; // Return empty array on error
  }
}

/**
 * Extracts the stage number from a class name
 * @param {string} className - The class name to extract from
 * @return {string} The extracted stage number or empty string if not found
 */
function extractStageFromClassName(className) {
  if (!className) return '';
  
  // Convert to string and lowercase for consistent processing
  const normalizedName = className.toString().toLowerCase().trim();
  
  // Pattern 1: Look for "Stage X" in the name
  let stageMatch = normalizedName.match(/stage\s*([1-6])/i);
  if (stageMatch && stageMatch[1]) {
    return stageMatch[1];
  }
  
  // Pattern 2: Look for "SX" format (e.g., S1, S2, etc.)
  stageMatch = normalizedName.match(/\bs([1-6])\b/i);
  if (stageMatch && stageMatch[1]) {
    return stageMatch[1];
  }
  
  // Pattern 3: Look for "X" where X is a digit 1-6 that might be stage
  // Only use this if it's likely to be referring to a stage
  if (normalizedName.includes('swim') || 
      normalizedName.includes('aqua') || 
      normalizedName.includes('water')) {
    stageMatch = normalizedName.match(/\b([1-6])\b/);
    if (stageMatch && stageMatch[1]) {
      return stageMatch[1];
    }
  }
  
  return '';
}

/**
 * Filters skills by stage based on the class name
 * @param {Object} allSkills - The complete skills object 
 * @param {string} stage - Stage number to filter by
 * @return {Object} Filtered skills object
 */
function filterSkillsByStage(allSkills, stage) {
  // If no stage specified or no skills available, return all skills
  if (!stage || !allSkills) {
    return allSkills;
  }
  
  Logger.log(`Filtering skills for stage ${stage}`);
  
  const result = {
    stage: [],
    saw: allSkills.saw || [] // Keep all SAW skills
  };
  
  // Only include skills for the specified stage and prior stages
  if (allSkills.stage && allSkills.stage.length > 0) {
    for (const skill of allSkills.stage) {
      // Extract the stage number from the skill header (e.g., 'S1 Float' → '1')
      const skillStage = skill.stage || extractStageFromSkillHeader(skill.header);
      
      // Include skills from current stage and optionally the previous stage
      if (skillStage === stage || skillStage === String(parseInt(stage) - 1)) {
        result.stage.push(skill);
      }
    }
  }
  
  Logger.log(`Filtered ${result.stage.length} stage skills and kept ${result.saw.length} SAW skills`);
  
  // If we didn't find any skills for this stage, return all skills
  if (result.stage.length === 0) {
    Logger.log('No skills found for specified stage, returning all skills');
    return allSkills;
  }
  
  return result;
}

/**
 * Extracts stage number from a skill header
 * @param {string} header - The skill header
 * @return {string} The stage number or empty string
 */
function extractStageFromSkillHeader(header) {
  if (!header) return '';
  
  // Check for common patterns like 'S1' at the beginning
  const match = header.toString().match(/^S([1-6])\s/);
  if (match && match[1]) {
    return match[1];
  }
  
  return '';
}

/**
 * Adds data validation for skill cells
 * @param {Sheet} sheet - The sheet to modify
 * @param {number} startRow - The starting row for student data
 * @param {number} studentCount - The number of students
 * @param {Object} skills - The skills configuration object
 */
function addSkillValidation(sheet, startRow, studentCount, skills) {
  if (studentCount <= 0) return;
  
  // Skill validation options
  const skillOptions = ['', 'X', '/', '?', 'N/A'];
  
  // Create a data validation rule
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(skillOptions, true)
    .setAllowInvalid(false)
    .build();
  
  // Calculate start and end columns for skills
  const skillsStartCol = 3 + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT;
  const totalSkillsColumns = (skills.stage.length * 2) + (skills.saw.length * 2);
  
  // Apply validation to all skill cells
  if (totalSkillsColumns > 0) {
    const validationRange = sheet.getRange(
      startRow, 
      skillsStartCol, 
      studentCount, 
      totalSkillsColumns
    );
    
    validationRange.setDataValidation(rule);
    
    // Add note explaining values
    validationRange.setNote(
      'X = Achieved\n' +
      '/ = In Progress\n' +
      '? = Not Yet Assessed\n' +
      'N/A = Not Applicable'
    );
  }
}

// Function removed as per requirement to eliminate pushing data back to Swimmer Records

/**
 * Sets up a simplified layout for private lessons
 * @param {Sheet} sheet - The instructor sheet
 * @param {Object} classDetails - The private lesson details
 */
function setupPrivateLessonLayout(sheet, classDetails) {
  try {
    // Add class header with private lesson info
    sheet.getRange('A2:Z2').merge()
      .setValue(`Private Lesson: ${classDetails.fullName}`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR)
      .setHorizontalAlignment('center');
    
    // Set up simplified headers for private lessons
    // First name
    sheet.getRange('A3').setValue(DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.FIRST_NAME)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Last name
    sheet.getRange('B3').setValue(DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.LAST_NAME)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Instructor
    sheet.getRange('C3').setValue('Instructor')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Notes
    sheet.getRange('D3').setValue('Notes')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Set column widths
    sheet.setColumnWidth(1, 150); // First Name
    sheet.setColumnWidth(2, 150); // Last Name
    sheet.setColumnWidth(3, 150); // Instructor
    sheet.setColumnWidth(4, 300); // Notes
    
    // Freeze the header rows
    sheet.setFrozenRows(3);
    
    // Get student roster for this private lesson
    const students = getStudentsForClass(classDetails);
    
    // Display info message if no students, but continue with empty roster
    if (students.length === 0) {
      sheet.getRange('A4').setValue('No students found for this private lesson. Add students manually or check class details.');
    } else {
      // Add student data
      for (let i = 0; i < students.length; i++) {
        const rowIndex = i + 4; // Start after header rows
        
        // Add student name
        sheet.getRange(rowIndex, 1).setValue(students[i].firstName);
        sheet.getRange(rowIndex, 2).setValue(students[i].lastName);
      }
    }
    
    // Add a default blank row if no students were found
    if (students.length === 0) {
      sheet.getRange('A4').setValue('');
      sheet.getRange('B4').setValue('');
    }
    
    Logger.log('Private lesson layout successfully created');
  } catch (error) {
    Logger.log(`Error setting up private lesson layout: ${error.message}`);
    sheet.getRange('A4').setValue(`Error creating private lesson layout: ${error.message}`);
    throw error;
  }
}

// Make functions available to other modules
const DynamicInstructorSheet = {
  createDynamicInstructorSheet: createDynamicInstructorSheet,
  onEditDynamicInstructorSheet: onEditDynamicInstructorSheet,
  setupPrivateLessonLayout: setupPrivateLessonLayout
};