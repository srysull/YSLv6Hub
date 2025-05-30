/**
 * YSL Hub Dynamic Instructor Sheet Module
 * 
 * This module implements a new approach to instructor sheets with a single dynamic
 * sheet that updates based on class selection. It pulls student data from rosters
 * and skills data from the Swimmer Records Workbook.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

// Configuration constants
const DYNAMIC_INSTRUCTOR_CONFIG = {
  SHEET_NAME: 'Instructor Sheet',
  ROSTER_SHEET_NAME: 'Daxko', // The sheet containing student registration data
  // Don't try to access the Excel file directly
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
    AFTER_COLUMN_COLOR: '#FFEBCC',   // Light orange for 'After' columns
    SELECTOR_BG_COLOR: '#D9EAD3'     // Light green for selector row
  },
  DAXKO_COLUMNS: {
    FIRST_NAME: 2, // Column C (0-indexed) - Student first name
    LAST_NAME: 3,  // Column D - Student last name
    PROGRAM: 22,   // Column W - Program name/description
    SESSION: 23,   // Column X - Session details (day/time)
    SESSION_DATE: 27 // Column AB - Session date
  },
  // Template structure constants
  TEMPLATE: {
    FROZEN_ROWS: 8, // Changed from 7 to 8 to accommodate selector row
    ROW_HEIGHTS: {
      SELECTOR_ROW: 30,
      HEADER_ROWS: 25
    }
  }
};

/**
 * Creates or resets the dynamic class hub sheet
 * @return {Sheet} The created or updated sheet
 */
function createDynamicInstructorSheet() { // Keeping the same function name but updating docs
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    
    // Create the sheet if it doesn't exist or completely reset it if it does
    if (!sheet) {
      // Create a new sheet
      sheet = ss.insertSheet(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    } else {
      // Clear everything from the existing sheet
      sheet.clear();
      sheet.clearFormats();
      sheet.clearConditionalFormatRules();
      
      // Clear data validations manually since clearDataValidations() might not be available
      try {
        // Try using clearDataValidations if available
        if (typeof sheet.clearDataValidations === 'function') {
          sheet.clearDataValidations();
        } else {
          // Fallback method: clear validation range by range
          const totalRows = sheet.getMaxRows();
          const totalCols = sheet.getMaxColumns();
          // Clear validations in smaller chunks to avoid timeout
          for (let startRow = 1; startRow <= totalRows; startRow += 20) {
            const rowsToProcess = Math.min(20, totalRows - startRow + 1);
            sheet.getRange(startRow, 1, rowsToProcess, totalCols).setDataValidation(null);
          }
        }
      } catch (e) {
        Logger.log(`Error clearing data validations: ${e.message}. Continuing anyway.`);
        // Continue with the rest of the initialization
      }
      
      // Reset all column widths to default
      const totalColumns = sheet.getMaxColumns();
      for (let i = 1; i <= totalColumns; i++) {
        sheet.setColumnWidth(i, 100); // Reset to default width
      }
      
      // Reset all row heights to default
      const totalRows = sheet.getMaxRows();
      for (let i = 1; i <= totalRows; i++) {
        sheet.setRowHeight(i, 21); // Reset to default height
      }
      
      // Ensure there are enough rows and columns
      const minRows = 100;
      const minColumns = 30;
      
      if (sheet.getMaxRows() < minRows) {
        sheet.insertRowsAfter(sheet.getMaxRows(), minRows - sheet.getMaxRows());
      }
      
      if (sheet.getMaxColumns() < minColumns) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), minColumns - sheet.getMaxColumns());
      }
      
      // Unhide any hidden rows or columns
      sheet.showRows(1, sheet.getMaxRows());
      sheet.showColumns(1, sheet.getMaxColumns());
    }
    
    // Only set up the selector row (row 1), not the full structure
    setupSheetSelector(sheet);
    
    // Create class selector dropdown
    createClassSelector(sheet);
    
    // Add onEdit trigger for the sheet
    ensureTriggerExists();
    
    // Clear any stored class selection
    PropertiesService.getDocumentProperties().deleteProperty('SELECTED_CLASS');
    
    // Show confirmation with clear instructions
    SpreadsheetApp.getUi().alert(
      'Instructor Sheet Created',
      'The instructor sheet has been created with a class selector. Please:\n\n' +
      '1. Select a class from the dropdown\n' +
      '2. Go to YSL Hub > Instructor Sheets > Update with Selected Class\n\n' +
      'This will populate the sheet with the template and student data.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Set active sheet to instructor sheet
    sheet.activate();
    
    return sheet;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'createDynamicInstructorSheet', 
        'Error creating instructor sheet. Please try again or contact support.');
    } else {
      Logger.log(`Error creating instructor sheet: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to create instructor sheet: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return null;
  }
}

/**
 * Sets up just the selector row of the instructor sheet
 * @param {Sheet} sheet - The sheet to set up
 */
function setupSheetSelector(sheet) {
  // Set row heights
  sheet.setRowHeight(1, DYNAMIC_INSTRUCTOR_CONFIG.TEMPLATE.ROW_HEIGHTS.SELECTOR_ROW); // Selector row
  
  // Set column widths
  sheet.setColumnWidth(1, 150); // Label
  sheet.setColumnWidth(2, 60); // Part of label
  sheet.setColumnWidth(3, 150); // Selector dropdown
  sheet.setColumnWidth(4, 150); // Selector dropdown extension
  
  // Set up selector row
  sheet.getRange('A1:B1').merge()
    .setValue(DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.CLASS_SELECTOR_LABEL)
    .setFontWeight('bold')
    .setHorizontalAlignment('right')
    .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SELECTOR_BG_COLOR);
  
  sheet.getRange('C1:D1').merge()
    .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SELECTOR_BG_COLOR); // Space for class selector dropdown
  
  // Add instructions in row 2
  sheet.getRange('A2:Z2').merge()
    .setValue('Select a class above, then use "YSL Hub > Instructor Sheets > Update with Selected Class" to generate the sheet')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Set frozen rows to just the selector row initially
  sheet.setFrozenRows(1);
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
    classRange.setNote('Select a class, then use "YSL Hub > Instructor Sheets > Update with Selected Class" to generate the sheet with this class.');
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
    
    // Also try to get Daxko sheet to extract private lessons
    const daxkoSheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME);
    
    if (!classesSheet && !daxkoSheet) {
      throw new Error('No source sheets found for class data');
    }
    
    const classNames = [];
    
    // First get regular class data if available
    if (classesSheet) {
      const classData = classesSheet.getDataRange().getValues();
      
      // Display some information about the Classes sheet
      Logger.log(`Classes sheet has ${classData.length} rows`);
      if (classData.length > 0) {
        Logger.log(`Classes sheet headers: ${JSON.stringify(classData[0])}`);
      }
      
      // Skip header row
      for (let i = 1; i < classData.length; i++) {
        // Check if row has valid data for the class definition
        if (classData[i].length > 3 && classData[i][1] && classData[i][2] && classData[i][3]) {
          const className = `${classData[i][1]} ${classData[i][2]} ${classData[i][3]}`;
          classNames.push(className);
          Logger.log(`Added regular class: ${className}`);
        }
      }
    }
    
    // Add private lessons from Daxko roster if available
    if (daxkoSheet) {
      const daxkoData = daxkoSheet.getDataRange().getValues();
      
      // Skip header row, process Daxko data
      if (daxkoData.length > 1) {
        const privateLessons = new Set(); // Use a Set to avoid duplicates
        
        for (let i = 1; i < daxkoData.length; i++) {
          // Check for sufficient data in this row
          if (daxkoData[i].length <= Math.max(
              DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM,
              DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION)) {
            continue;
          }
          
          const program = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM];
          let day = '';
          
          // Try to extract day from column Z if available (column 25, 0-indexed)
          const dayColumnIndex = 25; // Column Z (0-indexed)
          if (daxkoData[i].length > dayColumnIndex) {
            day = daxkoData[i][dayColumnIndex];
          }
          
          // Skip rows without program information
          if (!program) continue;
          
          // Check if this is a private lesson
          const programStr = program.toString().toLowerCase();
          if (programStr.includes('private')) {
            // Format as "[Program] [Day]" only for private lessons
            // Don't include the session time in the selector
            const privateLessonName = day ? `${program} ${day}` : program;
            privateLessons.add(privateLessonName);
          }
        }
        
        // Add private lessons to class names
        privateLessons.forEach(lesson => {
          classNames.push(lesson);
          Logger.log(`Added private lesson: ${lesson}`);
        });
      }
    }
    
    // If no classes found, add some test options
    if (classNames.length === 0) {
      Logger.log('No classes found, adding test classes');
      classNames.push('Test Swimming Monday 9:00 AM');
      classNames.push('Test Swimming Tuesday 10:00 AM');
      classNames.push('Private Swim Lesson Monday');
    }
    
    return classNames;
  } catch (error) {
    Logger.log(`Error getting available classes: ${error.message}`);
    
    // Return some test classes as a fallback
    return [
      'Test Swimming Monday 9:00 AM', 
      'Test Swimming Tuesday 10:00 AM',
      'Private Swim Lesson Monday'
    ];
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
      // Get the value directly from the cell since e.value might not be reliable for merged cells
      const selectedClass = e.range.getSheet().getRange('C1:D1').getValue();
      
      if (!selectedClass) {
        return; // No class selected
      }
      
      Logger.log(`Class selected from dropdown: ${selectedClass}`);
      
      // Store the selected class in the sheet properties so it can be referenced later when rebuilding
      PropertiesService.getDocumentProperties().setProperty('SELECTED_CLASS', selectedClass);
      
      // Show a toast message instructing the user to use the Rebuild option
      e.range.getSheet().getParent().toast(
        'Use "Rebuild Instructor Sheet" from the YSL Hub menu to apply this selection', 
        'Class Selected', 
        8 // Show for 8 seconds
      );
      
      // IMPORTANT: Return immediately to prevent the old function body from running
      return;
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
    // Get class details from the selected class
    const classDetails = parseClassDetails(selectedClass);
    
    // Check if this is a private lesson
    if (classDetails.isPrivateLesson) {
      // Use special layout for private lessons
      setupPrivateLessonLayout(sheet, classDetails);
      return; // Exit early, no need to proceed with regular class setup
    }
    
    // For regular classes, continue with the standard layout
    
    // Extract stage from class name if possible (e.g., "S1" from "Swimming S1 Monday")
    const stageInfo = extractStageFromClassName(classDetails.program);
    const stageDisplay = stageInfo.value ? `${stageInfo.prefix}${stageInfo.value.toUpperCase()}` : '';
    Logger.log(`Extracted stage: ${stageDisplay} from class: ${classDetails.program}`);
    
    // Add class header with stage info
    sheet.getRange('A2:Z2').merge()
      .setValue(`Class: ${selectedClass}${stageDisplay ? ` - Stage ${stageDisplay}` : ''}`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR)
      .setHorizontalAlignment('center');
    
    // Get student roster for this class first (we need this regardless of skills)
    const students = getStudentsForClass(classDetails);
    
    // Set up attendance columns - always needed
    setupAttendanceColumns(sheet);
      
    // Get skills - if this fails, we'll use fallbacks
    let skills;
    try {
      const allSkills = getSkillsFromSwimmerRecords();
      skills = filterSkillsByStage(allSkills, stageInfo);
    } catch (skillError) {
      Logger.log(`Skills data could not be loaded: ${skillError.message}`);
      
      // Use empty skills structure as fallback
      skills = {
        stage: [],
        saw: []
      };
      
      // Show a non-blocking notification to the user
      sheet.getRange('A3:B3').merge()
        .setValue('Note: Skills data unavailable. Only attendance tracking is enabled.')
        .setFontStyle('italic')
        .setFontColor('#CC0000');
    }
    
    // Always set up skill columns (may be empty if skills couldn't be loaded)
    setupSplitSkillsColumns(sheet, skills);
    
    // Always populate student data (even if no skills were found)
    populateStudentData(sheet, students, skills);
    
    // Reset frozen rows to show both header rows
    sheet.setFrozenRows(3);
    
  } catch (error) {
    Logger.log(`Error populating sheet with class data: ${error.message}`);
    
    // Create a minimal fallback layout
    sheet.getRange('A2').setValue(`Error loading class "${selectedClass}": ${error.message}`);
    sheet.getRange('A3').setValue('Please try again or contact support.');
    
    // Log extended error info
    console.error(`Failed to populate class data: ${error.stack || error.message}`);
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
    // First try to get the Daxko sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let daxkoSheet;
    
    try {
      daxkoSheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME);
      if (!daxkoSheet) {
        Logger.log(`${DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME} sheet not found. Will use test students.`);
        return createTestStudents(classDetails);
      }
    } catch (sheetError) {
      Logger.log(`Error accessing Daxko sheet: ${sheetError.message}`);
      return createTestStudents(classDetails);
    }
    
    // Get roster data with error handling
    let daxkoData;
    try {
      daxkoData = daxkoSheet.getDataRange().getValues();
      
      if (daxkoData.length <= 1) {
        Logger.log('No data found in Daxko sheet (only headers or empty)');
        return createTestStudents(classDetails);
      }
    } catch (dataError) {
      Logger.log(`Error reading Daxko data: ${dataError.message}`);
      return createTestStudents(classDetails);
    }
    
    // Log the class details for debugging
    Logger.log(`Looking for students in class: ${JSON.stringify(classDetails)}`);
    
    // Get the column headers to verify we're looking at the right columns
    const headers = daxkoData[0];
    
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
      const stageInfo = extractStageFromClassName(classDetails.program);
      if (stageInfo.value) {
        searchTerms.push(`stage ${stageInfo.value}`);
        searchTerms.push(`${stageInfo.prefix}${stageInfo.value}`);
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
    let matchCount = 0;
    
    // Direct matching on program and session columns
    for (let i = 1; i < daxkoData.length; i++) {
      try {
        // Basic validation first
        if (daxkoData[i].length <= Math.max(
            DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.FIRST_NAME,
            DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.LAST_NAME)) {
          continue; // Skip rows with insufficient data
        }
        
        const firstName = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.FIRST_NAME];
        const lastName = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.LAST_NAME];
        
        // Skip rows without names
        if (!firstName || !lastName) continue;
        
        // Get all columns that might contain class information (with error handling)
        let allFieldsStr = '';
        
        try {
          // Get program and session info
          const rowProgram = daxkoData[i].length > DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM 
              ? daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM] || '' 
              : '';
          
          const rowSession = daxkoData[i].length > DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION 
              ? daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION] || '' 
              : '';
          
          // Convert to strings for consistent comparison
          const programStr = rowProgram ? rowProgram.toString().toLowerCase().trim() : '';
          const sessionStr = rowSession ? rowSession.toString().toLowerCase().trim() : '';
          
          // Build the search string
          allFieldsStr = `${programStr} ${sessionStr}`;
        } catch (columnError) {
          Logger.log(`Error processing column data: ${columnError.message}`);
          continue; // Skip this row
        }
        
        // Enhanced matching logic
        let isMatch = false;
        let matchReason = '';
        
        // Count how many search terms match this student record
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
        
        // If we match enough search terms, consider it a match
        if (totalTerms > 0 && termMatches >= Math.ceil(totalTerms * 0.5)) {
          isMatch = true;
          matchReason = `Matched ${termMatches}/${totalTerms} search terms`;
        }
        
        // Class-specific matching for full class name
        if (!isMatch && classDetails.fullName) {
          const normalizedClassName = classDetails.fullName.toLowerCase().trim();
          
          if (allFieldsStr.includes(normalizedClassName)) {
            isMatch = true;
            matchReason = 'Matched full class name';
          }
          else if (normalizedClassName.includes('test') || allFieldsStr.includes('test')) {
            isMatch = true;
            matchReason = 'Test class match';
          }
        }
        
        if (isMatch) {
          matchCount++;
          students.push({
            firstName: firstName,
            lastName: lastName,
            fullName: `${firstName} ${lastName}`,
            skills: {}, // Will be populated later
            matchReason: matchReason // For debugging
          });
          
          // Log sample of matched students
          if (matchCount <= 3) {
            Logger.log(`Found matching student: ${firstName} ${lastName} (${matchReason})`);
          }
        }
      } catch (rowError) {
        // If anything goes wrong processing this row, skip it and continue
        Logger.log(`Error processing student row ${i}: ${rowError.message}`);
        continue;
      }
    }
    
    // Log results
    Logger.log(`Found ${students.length} matching students for class ${classDetails.fullName}`);
    
    // If no students found, create test students
    if (students.length === 0) {
      Logger.log('No students found. Creating test students as fallback.');
      return createTestStudents(classDetails);
    }
    
    return students;
    
  } catch (error) {
    // If any error occurs, log it and return test students instead of throwing
    Logger.log(`Error getting students for class: ${error.message}`);
    return createTestStudents(classDetails);
  }
}

/**
 * Creates test students when real students can't be found
 * @param {Object} classDetails - The class details
 * @return {Array} Array of test student objects
 */
function createTestStudents(classDetails) {
  try {
    // Create test students based on class details
    const stageInfo = extractStageFromClassName(classDetails.program);
    const stageCode = stageInfo.value ? `${stageInfo.prefix}${stageInfo.value}` : '';
    
    // Create several test students for better testing
    const testStudents = [];
    
    // Basic test students
    testStudents.push({
      firstName: 'Test',
      lastName: stageCode ? `Student ${stageCode}` : 'Student1',
      fullName: stageCode ? `Test Student ${stageCode}` : 'Test Student1',
      skills: {},
      matchReason: 'Test student'
    });
    
    testStudents.push({
      firstName: 'Test',
      lastName: stageCode ? `Student ${stageCode}-2` : 'Student2',
      fullName: stageCode ? `Test Student ${stageCode}-2` : 'Test Student2',
      skills: {},
      matchReason: 'Test student'
    });
    
    // Add a few more test students with random names for better testing
    const firstNames = ['Alex', 'Jordan', 'Casey', 'Morgan', 'Taylor'];
    const lastNames = ['Smith', 'Johnson', 'Williams', 'Brown', 'Jones'];
    
    for (let i = 0; i < 3; i++) {
      const firstName = firstNames[Math.floor(Math.random() * firstNames.length)];
      const lastName = lastNames[Math.floor(Math.random() * lastNames.length)];
      
      testStudents.push({
        firstName: firstName,
        lastName: lastName,
        fullName: `${firstName} ${lastName}`,
        skills: {},
        matchReason: 'Random test student'
      });
    }
    
    Logger.log(`Created ${testStudents.length} test students for ${classDetails.fullName}`);
    return testStudents;
    
  } catch (error) {
    // If even creating test students fails, return a minimal set
    Logger.log(`Error creating test students: ${error.message}`);
    return [
      { firstName: 'Test', lastName: 'Student1', fullName: 'Test Student1', skills: {}, matchReason: 'Fallback test student' },
      { firstName: 'Test', lastName: 'Student2', fullName: 'Test Student2', skills: {}, matchReason: 'Fallback test student' }
    ];
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
    // Get Swimmer Records URL from different possible sources with robust error handling
    let swimmerRecordsUrl = null;
    
    // Try different ways to get the URL
    try {
      // First, try GlobalFunctions if available
      if (typeof GlobalFunctions !== 'undefined' && typeof GlobalFunctions.safeGetProperty === 'function') {
        swimmerRecordsUrl = GlobalFunctions.safeGetProperty('swimmerRecordsUrl');
        
        // If not found with direct property name, try through CONFIG object
        if (!swimmerRecordsUrl && typeof CONFIG !== 'undefined' && CONFIG.SWIMMER_RECORDS_URL) {
          swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL);
        }
      }
      
      // If still not found, try getting from AdministrativeModule
      if (!swimmerRecordsUrl && typeof AdministrativeModule !== 'undefined' && 
          typeof AdministrativeModule.getSystemConfiguration === 'function') {
        const config = AdministrativeModule.getSystemConfiguration();
        if (config && config.swimmerRecordsUrl) {
          swimmerRecordsUrl = config.swimmerRecordsUrl;
        }
      }
      
      // Last resort - direct property access
      if (!swimmerRecordsUrl) {
        swimmerRecordsUrl = PropertiesService.getScriptProperties().getProperty('swimmerRecordsUrl');
      }
    } catch (propError) {
      Logger.log(`Error accessing configuration: ${propError.message}`);
      // Don't return yet, continue with hardcoded ID as last resort
    }
    
    // If URL not found from any source, try a known hardcoded ID for testing
    if (!swimmerRecordsUrl) {
      Logger.log('Swimmer Records URL not found in system configuration. Using fallback skills.');
      return createFallbackSkills();
    }
    
    // Extract spreadsheet ID from URL with multiple fallback methods
    let ssId = null;
    
    try {
      // Try with GlobalFunctions first
      if (typeof GlobalFunctions !== 'undefined' && typeof GlobalFunctions.extractIdFromUrl === 'function') {
        ssId = GlobalFunctions.extractIdFromUrl(swimmerRecordsUrl);
      }
      
      // If that fails, try direct regex extraction
      if (!ssId) {
        const urlPattern = /[-\w]{25,}/;
        const match = swimmerRecordsUrl.match(urlPattern);
        ssId = match ? match[0] : null;
      }
      
      // If that fails too, just use the URL as-is
      if (!ssId) {
        ssId = swimmerRecordsUrl;
      }
    } catch (urlError) {
      Logger.log(`Error extracting ID from URL: ${urlError.message}`);
      // Try using URL directly
      ssId = swimmerRecordsUrl;
    }
    
    if (!ssId) {
      Logger.log('Invalid Swimmer Records URL - could not extract valid ID');
      return createFallbackSkills();
    }
    
    // Log the spreadsheet ID we're trying to open
    Logger.log(`Attempting to open Swimmer Records with ID: ${ssId}`);
    
    try {
      // Try to open the Swimmer Records Workbook with careful error handling
      let swimmerSS = null;
      
      try {
        // Try using GlobalFunctions for safer access if available
        if (typeof GlobalFunctions !== 'undefined' && typeof GlobalFunctions.safeGetSpreadsheetById === 'function') {
          swimmerSS = GlobalFunctions.safeGetSpreadsheetById(ssId);
        } else {
          // Direct access as fallback
          swimmerSS = SpreadsheetApp.openById(ssId);
        }
      } catch (accessError) {
        const errorMsg = `Error accessing Swimmer Records: ${accessError.message}`;
        Logger.log(errorMsg);
        
        // Log with ErrorHandling if available
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(errorMsg, 'ERROR', 'getSkillsFromSwimmerRecords');
        }
        
        // Fall back to test data
        return createFallbackSkills();
      }
      
      if (!swimmerSS) {
        Logger.log('Could not open Swimmer Records spreadsheet - null result');
        return createFallbackSkills();
      }
      
      // Find the right sheet - either Skills sheet or first sheet
      let swimmerSheet = null;
      try {
        // Try to get Skills sheet first
        swimmerSheet = swimmerSS.getSheetByName('Skills');
        
        // If not found, try first sheet
        if (!swimmerSheet) {
          const sheets = swimmerSS.getSheets();
          if (sheets && sheets.length > 0) {
            swimmerSheet = sheets[0];
          }
        }
      } catch (sheetError) {
        Logger.log(`Error finding appropriate sheet: ${sheetError.message}`);
        return createFallbackSkills();
      }
      
      if (!swimmerSheet) {
        Logger.log('No suitable sheet found in Swimmer Records Workbook');
        return createFallbackSkills();
      }
      
      // Get the header row with error handling
      let headerRow = null;
      try {
        headerRow = swimmerSheet.getRange(1, 1, 1, swimmerSheet.getLastColumn()).getValues()[0];
      } catch (rangeError) {
        Logger.log(`Error reading header row: ${rangeError.message}`);
        return createFallbackSkills();
      }
      
      if (!headerRow || headerRow.length === 0) {
        Logger.log('Empty header row in Swimmer Records workbook');
        return createFallbackSkills();
      }
      
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
      
    } catch (finalError) {
      const errorMsg = `Failed to process Swimmer Records: ${finalError.message}`;
      Logger.log(errorMsg);
      
      // Log with ErrorHandling if available
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage(errorMsg, 'ERROR', 'getSkillsFromSwimmerRecords');
      }
      
      return createFallbackSkills();
    }
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
 * Extracts the stage from a class name
 * @param {string} className - The class name to extract from
 * @return {Object} Contains stage value and prefix (e.g. {value: '1', prefix: 'S'})
 */
function extractStageFromClassName(className) {
  if (!className) return { value: '', prefix: '' };
  
  // Convert to string and lowercase for consistent processing
  const normalizedName = className.toString().toLowerCase().trim();
  
  // Pattern 1: Look for "Stage X" in the name (X can be digit or letter)
  let stageMatch = normalizedName.match(/stage\s*([1-6a-f])/i);
  if (stageMatch && stageMatch[1]) {
    // Determine if it's a numeric or letter stage
    const stageValue = stageMatch[1];
    return { 
      value: stageValue,
      prefix: 'S'  // Default prefix for all stages
    };
  }
  
  // Pattern 2: Look for "SX" format (e.g., S1, S2, SA, etc.)
  stageMatch = normalizedName.match(/\bs([1-6a-f])\b/i);
  if (stageMatch && stageMatch[1]) {
    return { 
      value: stageMatch[1],
      prefix: 'S'
    };
  }
  
  // Pattern 3: Look for "X" where X is a digit 1-6 or letter A-F that might be stage
  // Only use this if it's likely to be referring to a stage
  if (normalizedName.includes('swim') || 
      normalizedName.includes('aqua') || 
      normalizedName.includes('water')) {
    stageMatch = normalizedName.match(/\b([1-6a-f])\b/i);
    if (stageMatch && stageMatch[1]) {
      return { 
        value: stageMatch[1],
        prefix: 'S'
      };
    }
  }
  
  // For backward compatibility, return an empty string when no stage is found
  return { value: '', prefix: '' };
}

/**
 * Filters skills by stage based on the class name
 * @param {Object} allSkills - The complete skills object 
 * @param {Object} stageInfo - Stage info with value and prefix
 * @return {Object} Filtered skills object
 */
function filterSkillsByStage(allSkills, stageInfo) {
  // If no stage specified or no skills available, return all skills
  if (!stageInfo || !stageInfo.value || !allSkills) {
    return allSkills;
  }
  
  const stageValue = stageInfo.value.toString().toLowerCase();
  const stagePrefix = stageInfo.prefix || 'S';
  const stageCode = `${stagePrefix}${stageValue}`;
  
  Logger.log(`Filtering skills for stage ${stageCode}`);
  
  const result = {
    stage: [],
    saw: allSkills.saw || [] // Keep all SAW skills
  };
  
  // Only include skills for the specified stage and prior stages
  if (allSkills.stage && allSkills.stage.length > 0) {
    for (const skill of allSkills.stage) {
      // Extract the stage from the skill header (e.g., 'S1 Float' → 'S1')
      const skillStageInfo = extractStageFromSkillHeader(skill.header);
      
      // For numeric stages, include current stage and previous stage
      if (/^[0-9]+$/.test(stageValue)) {
        const prevStage = String(parseInt(stageValue) - 1);
        
        if (skillStageInfo === stageCode || 
            skillStageInfo === `${stagePrefix}${prevStage}`) {
          result.stage.push(skill);
        }
      } 
      // For letter stages (like 'A'), only include exact matches
      else {
        if (skillStageInfo === stageCode) {
          result.stage.push(skill);
        }
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
 * Extracts stage code from a skill header
 * @param {string} header - The skill header
 * @return {string} The complete stage code (e.g., 'S1', 'SA') or empty string
 */
function extractStageFromSkillHeader(header) {
  if (!header) return '';
  
  // Check for common patterns like 'S1' or 'SA' at the beginning
  const match = header.toString().match(/^(S[1-6A-Fa-f])\s/);
  if (match && match[1]) {
    return match[1].toUpperCase();
  }
  
  return '';
}

/**
 * Formats skill cells without adding data validation (using simple text format)
 * @param {Sheet} sheet - The sheet to modify
 * @param {number} startRow - The starting row for student data
 * @param {number} studentCount - The number of students
 * @param {Object} skills - The skills configuration object
 */
function addSkillValidation(sheet, startRow, studentCount, skills) {
  if (studentCount <= 0) return;
  
  try {
    // Check if we can access the header
    let headerText = '';
    try {
      headerText = sheet.getRange('A2').getValue().toString();
    } catch (e) {
      Logger.log(`Error accessing header: ${e.message}`);
      // Continue but assume it's not a private lesson
      headerText = '';
    }
    
    if (headerText.indexOf('Private Lesson:') >= 0) {
      // Don't format for private lessons
      return;
    }
    
    // Calculate start and end columns for skills
    const skillsStartCol = 3 + DYNAMIC_INSTRUCTOR_CONFIG.HEADERS.ATTENDANCE_COUNT;
    const totalSkillsColumns = (skills.stage.length * 2) + (skills.saw.length * 2);
    
    // Format the skill cells (but don't add validation dropdown)
    if (totalSkillsColumns > 0) {
      const skillCellsRange = sheet.getRange(
        startRow, 
        skillsStartCol, 
        studentCount, 
        totalSkillsColumns
      );
      
      // Set formatting for the skills cells (center alignment, no validation)
      skillCellsRange.setHorizontalAlignment('center');
      skillCellsRange.setVerticalAlignment('middle');
      
      // Add note explaining values to the header row to guide instructors
      const headerNote = 'Suggested values for skill cells:\n' +
                        'X = Achieved\n' +
                        '/ = In Progress\n' +
                        '? = Not Yet Assessed\n' +
                        'N/A = Not Applicable';
      
      // Add the note to the skills section headers
      sheet.getRange(3, skillsStartCol).setNote(headerNote);
      
      // If there are SAW skills, add the note to that header too
      if (skills.saw.length > 0) {
        const sawStartCol = skillsStartCol + (skills.stage.length * 2);
        sheet.getRange(3, sawStartCol).setNote(headerNote);
      }
    }
    
    // Hide any unused columns to clean up the sheet
    const requiredColumns = skillsStartCol + totalSkillsColumns;
    if (sheet.getMaxColumns() > requiredColumns + 1) { // +1 for buffer
      sheet.hideColumns(requiredColumns + 1, sheet.getMaxColumns() - requiredColumns);
    }
    
    // Hide any unused rows to clean up the sheet
    const requiredRows = startRow + studentCount;
    if (sheet.getMaxRows() > requiredRows + 3) { // +3 for buffer
      sheet.hideRows(requiredRows + 1, sheet.getMaxRows() - requiredRows);
    }
  } catch (error) {
    // Log the error but don't fail the whole function
    Logger.log(`Error formatting skill cells: ${error.message}`);
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
    // First clear any existing data validation to avoid errors
    try {
      const fullSheetRange = sheet.getDataRange();
      // Try using clearDataValidations if available
      if (typeof fullSheetRange.clearDataValidations === 'function') {
        fullSheetRange.clearDataValidations();
      } else {
        // Fallback: clear validation by setting to null
        fullSheetRange.setDataValidation(null);
      }
    } catch (validationError) {
      Logger.log(`Error clearing data validations: ${validationError.message}. Continuing anyway.`);
      // Continue with the setup process
    }
    
    // Clear all content except the class selector in row 1
    clearStudentData(sheet);
    
    // Get all students and lesson dates from Daxko for this private lesson
    const studentsWithDates = getPrivateLessonStudentsWithDates(classDetails);
    
    // Add class info in row 2 (don't merge to preserve layout)
    sheet.getRange('A2').setValue('Private Lesson:')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR);
    
    sheet.getRange('B2').setValue(classDetails.fullName)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR);
    
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
    
    // Session column (time details from Daxko column X)
    sheet.getRange('D3').setValue('Session')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Notes
    sheet.getRange('E3').setValue('Notes')
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Set column widths
    sheet.setColumnWidth(1, 150); // First Name
    sheet.setColumnWidth(2, 150); // Last Name
    sheet.setColumnWidth(3, 150); // Instructor
    sheet.setColumnWidth(4, 200); // Session
    sheet.setColumnWidth(5, 300); // Notes
    
    // Hide unused columns
    if (sheet.getMaxColumns() > 6) {
      sheet.hideColumns(6, sheet.getMaxColumns() - 5);
    }
    
    // Freeze the header rows
    sheet.setFrozenRows(3);
    
    // Calculate actual rows needed - use actual students found or minimum 5
    const actualRowsNeeded = Math.max(5, studentsWithDates.length);
    
    // Format all data columns to be centered
    const dataRowsRange = sheet.getRange(4, 1, actualRowsNeeded, 5);
    dataRowsRange.setHorizontalAlignment('center');
    dataRowsRange.setVerticalAlignment('middle');
    
    // Add student data if available
    if (studentsWithDates.length === 0) {
      // Add empty rows for manual entry
      for (let i = 0; i < actualRowsNeeded; i++) {
        const rowIndex = i + 4; // Start after header rows
        sheet.getRange(rowIndex, 1).setValue('');
        sheet.getRange(rowIndex, 2).setValue('');
        sheet.getRange(rowIndex, 3).setValue(''); // Instructor
        sheet.getRange(rowIndex, 4).setValue(''); // Session
      }
    } else {
      // Add existing students with dates and empty rows
      for (let i = 0; i < actualRowsNeeded; i++) {
        const rowIndex = i + 4; // Start after header rows
        
        if (i < studentsWithDates.length) {
          // Add student name and session info (from Daxko column X)
          sheet.getRange(rowIndex, 1).setValue(studentsWithDates[i].firstName);
          sheet.getRange(rowIndex, 2).setValue(studentsWithDates[i].lastName);
          // Leave instructor column blank for instructor input
          sheet.getRange(rowIndex, 3).setValue('');
          // Put session information in column 4
          sheet.getRange(rowIndex, 4).setValue(studentsWithDates[i].session || '');
        } else {
          // Add empty row for manual entry
          sheet.getRange(rowIndex, 1).setValue('');
          sheet.getRange(rowIndex, 2).setValue('');
          sheet.getRange(rowIndex, 3).setValue('');
          sheet.getRange(rowIndex, 4).setValue('');
        }
      }
    }
    
    // Add alternating row colors for better readability
    for (let i = 0; i < actualRowsNeeded; i++) {
      if (i % 2 === 1) { // Odd rows (0-based index, so rows 5, 7, 9, etc.)
        sheet.getRange(i + 4, 1, 1, 5).setBackground('#f3f3f3'); // Light gray
      }
    }
    
    // Hide any unused rows
    const totalRowsOnSheet = sheet.getMaxRows();
    if (totalRowsOnSheet > actualRowsNeeded + 3) { // +3 for header rows
      sheet.hideRows(actualRowsNeeded + 4, totalRowsOnSheet - (actualRowsNeeded + 3));
    }
    
    Logger.log('Private lesson layout successfully created');
  } catch (error) {
    Logger.log(`Error setting up private lesson layout: ${error.message}`);
    sheet.getRange('A4').setValue(`Error creating private lesson layout: ${error.message}`);
    throw error;
  }
}

/**
 * Gets private lesson students with their lesson dates, sorted by date (soonest first)
 * @param {Object} classDetails - The private lesson details
 * @return {Array} Array of student objects with dates
 */
function getPrivateLessonStudentsWithDates(classDetails) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const daxkoSheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME);
    
    if (!daxkoSheet) {
      throw new Error(`${DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME} sheet not found. Please make sure the Daxko sheet exists.`);
    }
    
    // Get all roster data from Daxko sheet
    const daxkoData = daxkoSheet.getDataRange().getValues();
    const studentsWithDates = [];
    
    // Process Daxko data to find private lesson students with dates
    for (let i = 1; i < daxkoData.length; i++) {
      // Check for sufficient data in this row
      if (daxkoData[i].length <= Math.max(
          DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.FIRST_NAME,
          DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.LAST_NAME,
          DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM,
          DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION,
          DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION_DATE)) {
        continue;
      }
      
      const firstName = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.FIRST_NAME];
      const lastName = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.LAST_NAME];
      const rowProgram = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.PROGRAM];
      const rowSession = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION];
      const rowDate = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION_DATE];
      
      // Skip rows without names or dates
      if (!firstName || !lastName) continue;
      
      // Normalize for comparison
      const normalizedClassName = classDetails.fullName.toLowerCase().trim();
      const normalizedProgram = rowProgram ? rowProgram.toString().toLowerCase().trim() : '';
      const normalizedSession = rowSession ? rowSession.toString().toLowerCase().trim() : '';
      
      // Check if this student is in the private lesson
      if ((normalizedProgram && normalizedClassName.includes(normalizedProgram)) ||
          (normalizedSession && normalizedClassName.includes(normalizedSession)) ||
          (normalizedProgram && normalizedProgram.includes('private'))) {
        
        // Format date for display - convert date object to string if needed
        let formattedDate = '';
        if (rowDate) {
          if (rowDate instanceof Date) {
            formattedDate = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
          } else {
            formattedDate = rowDate.toString();
          }
        }
        
        // Also get the session information (time details) from column X
        let sessionInfo = '';
        if (daxkoData[i].length > DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION) {
          sessionInfo = daxkoData[i][DYNAMIC_INSTRUCTOR_CONFIG.DAXKO_COLUMNS.SESSION];
          if (sessionInfo) {
            sessionInfo = sessionInfo.toString().trim();
          }
        }
        
        studentsWithDates.push({
          firstName: firstName,
          lastName: lastName,
          date: formattedDate,
          session: sessionInfo, // Add the session info
          // Store the original date for sorting
          rawDate: rowDate instanceof Date ? rowDate : null
        });
      }
    }
    
    // Sort students by date (soonest first)
    studentsWithDates.sort(function(a, b) {
      // If we have raw dates, use them for sorting
      if (a.rawDate && b.rawDate) {
        return a.rawDate - b.rawDate;
      }
      
      // Fall back to string comparison if raw dates not available
      if (a.date && b.date) {
        return a.date.localeCompare(b.date);
      }
      
      // Put entries with dates before entries without dates
      if (a.date && !b.date) return -1;
      if (!a.date && b.date) return 1;
      
      // If no dates available, sort by name
      return (a.firstName + a.lastName).localeCompare(b.firstName + b.lastName);
    });
    
    Logger.log(`Found ${studentsWithDates.length} private lesson students with dates`);
    
    // If no students found, create a fallback student
    if (studentsWithDates.length === 0) {
      studentsWithDates.push({
        firstName: 'Private',
        lastName: 'Student',
        date: ''
      });
    }
    
    return studentsWithDates;
  } catch (error) {
    Logger.log(`Error getting private lesson students: ${error.message}`);
    // Return a test student as fallback
    return [{
      firstName: 'Test',
      lastName: 'Student',
      date: ''
    }];
  }
}

/**
 * Rebuilds the instructor sheet based on the selected class
 * This is called from the menu item rather than automatically on edit
 */
/**
 * Rebuilds the dynamic instructor sheet with the selected class
 * Provides robust error handling to ensure the sheet is always usable
 * even if some data can't be loaded
 * 
 * @return {Sheet} The rebuilt sheet or null on failure
 */
function rebuildDynamicInstructorSheet() {
  try {
    // First get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    
    if (!sheet) {
      // If sheet doesn't exist, create it first
      return createDynamicInstructorSheet();
    }
    
    // Get the selected class from properties
    const selectedClass = PropertiesService.getDocumentProperties().getProperty('SELECTED_CLASS');
    
    // If no class is selected, get it from the sheet
    const classFromSheet = sheet.getRange('C1:D1').getValue();
    const classToUse = selectedClass || classFromSheet;
    
    if (!classToUse) {
      SpreadsheetApp.getUi().alert(
        'No Class Selected',
        'Please select a class from the dropdown before rebuilding the sheet.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return null;
    }
    
    // Store the class selection in case it came from the sheet
    if (!selectedClass && classFromSheet) {
      PropertiesService.getDocumentProperties().setProperty('SELECTED_CLASS', classFromSheet);
    }
    
    try {
      // Log what we're doing
      Logger.log(`Rebuilding instructor sheet for class: ${classToUse}`);
      
      // Preserve row 1 (the selector row)
      const selectorRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues();
      const selectorRowFormatting = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getBackgrounds();
      
      // Clear everything except row 1
      clearSheetExceptSelector(sheet);
      
      // Apply the template structure - this should never fail
      try {
        applyTemplateStructure(sheet, classToUse);
      } catch (templateError) {
        Logger.log(`Error applying template: ${templateError.message}, continuing with minimal structure`);
        // Create a minimal emergency header 
        sheet.getRange('A2').setValue(`Class: ${classToUse}`).setFontWeight('bold');
        sheet.getRange('A3').setValue('First Name').setFontWeight('bold');
        sheet.getRange('B3').setValue('Last Name').setFontWeight('bold');
      }
      
      // Restore the selector in row 1 - this should also never fail
      sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setValues(selectorRow);
      sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setBackgrounds(selectorRowFormatting);
      
      // Make sure the class selector dropdown is preserved
      try {
        createClassSelector(sheet);
      } catch (selectorError) {
        Logger.log(`Error recreating selector: ${selectorError.message}, continuing with static value`);
        // Just set the value without the dropdown validation
        sheet.getRange('C1').setValue(classToUse);
      }
      
      // Set the frozen rows to 8 (1 selector row + 7 template header rows) or default to 3 if that fails
      try {
        sheet.setFrozenRows(DYNAMIC_INSTRUCTOR_CONFIG.TEMPLATE.FROZEN_ROWS);
      } catch (e) {
        sheet.setFrozenRows(3); // Default to minimal frozen rows
      }
      
      // Parse the class details - this should be robust
      const classDetails = parseClassDetails(classToUse);
      
      // Populate with student data - this might fail due to external data access
      try {
        populateTemplateWithStudentData(sheet, classDetails);
      } catch (dataError) {
        Logger.log(`Error populating student data: ${dataError.message}, creating minimal empty structure`);
        
        // Add error message but keep the sheet usable
        sheet.getRange('A4').setValue('Data Error: Could not load student data.');
        sheet.getRange('A5').setValue(`Error message: ${dataError.message}`);
        sheet.getRange('A6').setValue('Please check system configuration and try again.');
        
        // Add a few empty rows for manual data entry
        sheet.getRange('A8').setValue('Student First Name');
        sheet.getRange('B8').setValue('Student Last Name');
        sheet.getRange('C8').setValue('Notes');
      }
      
      // Always show confirmation - the sheet was rebuilt even if some data is missing
      SpreadsheetApp.getUi().alert(
        'Instructor Sheet Rebuilt',
        `The instructor sheet has been rebuilt for class "${classToUse}"`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      // Set active sheet to instructor sheet
      sheet.activate();
      
      return sheet;
    } catch (mainError) {
      // Handle errors in the main process, but keep going
      Logger.log(`Error in main rebuild process: ${mainError.message}`);
      
      // Try to create a minimal usable sheet
      sheet.getRange('A2').setValue(`Class: ${classToUse}`).setFontWeight('bold');
      sheet.getRange('A3').setValue('Error creating instructor sheet. Please try again.');
      sheet.getRange('A4').setValue(`Error: ${mainError.message}`);
      
      SpreadsheetApp.getUi().alert(
        'Partial Sheet Created',
        `There was a problem creating the full instructor sheet, but a simple outline has been created. Error: ${mainError.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      return sheet;
    }
  } catch (error) {
    // Handle critical errors that prevent any sheet creation
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'rebuildDynamicInstructorSheet', 
        'Error rebuilding instructor sheet. Please try again or contact support.');
    } else {
      Logger.log(`Critical error rebuilding instructor sheet: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to rebuild instructor sheet: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return null;
  }
}

/**
 * Clears the sheet except for the selector row (row 1)
 * @param {Sheet} sheet - The sheet to clear
 */
function clearSheetExceptSelector(sheet) {
  const lastRow = sheet.getLastRow();
  const totalColumns = sheet.getMaxColumns();
  
  if (lastRow > 1) {
    // Clear contents, formats, and validations below row 1
    sheet.getRange(2, 1, lastRow - 1, totalColumns).clear();
    sheet.getRange(2, 1, lastRow - 1, totalColumns).clearFormat();
    
    // Try to clear data validations if available
    try {
      if (typeof sheet.getRange(2, 1, lastRow - 1, totalColumns).clearDataValidations === 'function') {
        sheet.getRange(2, 1, lastRow - 1, totalColumns).clearDataValidations();
      }
    } catch (e) {
      Logger.log(`Error clearing data validations: ${e.message}. Continuing anyway.`);
    }
  }
  
  // Reset column widths except row 1
  for (let i = 1; i <= totalColumns; i++) {
    sheet.setColumnWidth(i, 100); // Reset to default width
  }
  
  // Reset row heights except row 1
  for (let i = 2; i <= Math.max(lastRow, 100); i++) {
    sheet.setRowHeight(i, 21); // Reset to default height
  }
}

/**
 * Applies the template structure from the Excel template
 * @param {Sheet} sheet - The sheet to update
 * @param {string} className - The selected class name
 */
function applyTemplateStructure(sheet, className) {
  try {
    // Manual approach instead of trying to use the Excel file directly
    
    // Set the title in row 2
    sheet.getRange('A2:Z2').merge()
      .setValue('YSL v6 & SAW ASSESSMENT SHEET')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR);
    
    // Set the class information header in row 3
    sheet.getRange('A3:Z3').merge()
      .setValue('Class Information')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
    
    // Extract class details from the selected class
    const classDetails = parseClassDetails(className);
    
    // Set class details in rows 4-6
    // Row 4: Program & Day
    sheet.getRange('A4').setValue('Program:').setFontWeight('bold');
    sheet.getRange('B4:C4').merge().setValue(classDetails.program);
    sheet.getRange('D4').setValue('Day:').setFontWeight('bold');
    sheet.getRange('E4:F4').merge().setValue(classDetails.day);
    
    // Row 5: Location & Instructor & Time
    sheet.getRange('A5').setValue('Location:').setFontWeight('bold');
    sheet.getRange('B5:C5').merge().setValue('PenBay YMCA'); // Default location
    sheet.getRange('D5').setValue('Instructor:').setFontWeight('bold');
    sheet.getRange('E5:F5').merge(); // Empty for instructor to fill in
    sheet.getRange('G5').setValue('Time:').setFontWeight('bold'); 
    sheet.getRange('H5:I5').merge().setValue(classDetails.time); // Add time
    
    // Row 6: Students count
    sheet.getRange('A6').setValue('Students:').setFontWeight('bold');
    sheet.getRange('B6:C6').merge(); // Will be filled with student count later
    
    // Row 7: Empty spacing row
    
    // Row 8: Column headers for student assessment
    setupStudentColumnHeaders(sheet, classDetails);
    
    // Format all headers consistently
    sheet.getRange('A2:Z8').setHorizontalAlignment('center').setVerticalAlignment('middle');
    
    // Set row heights for consistent layout
    for (let i = 2; i <= 8; i++) {
      sheet.setRowHeight(i, DYNAMIC_INSTRUCTOR_CONFIG.TEMPLATE.ROW_HEIGHTS.HEADER_ROWS);
    }
  } catch (error) {
    Logger.log(`Error applying template structure: ${error.message}`);
    throw error;
  }
}

/**
 * Sets up student column headers in row 8 based on the template
 * @param {Sheet} sheet - The sheet to update
 * @param {Object} classDetails - The parsed class details
 */
function setupStudentColumnHeaders(sheet, classDetails) {
  // Row 8 contains student column headers
  sheet.getRange('A8').setValue('Date').setFontWeight('bold');
  
  // Columns B-F for assessment criteria labels
  sheet.getRange('B8:F8').merge().setValue('Assessment Criteria').setFontWeight('bold');
  
  // Format the header row
  sheet.getRange('A8:F8')
    .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
    .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR);
  
  // Set column widths for assessment criteria section
  sheet.setColumnWidth(1, 100); // Date
  sheet.setColumnWidth(2, 120); // Assessment criteria
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 120);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 120);
}

/**
 * Populates the template with student data based on class details
 * @param {Sheet} sheet - The sheet to update
 * @param {Object} classDetails - The parsed class details
 */
function populateTemplateWithStudentData(sheet, classDetails) {
  try {
    // Get students for this class
    const students = getStudentsForClass(classDetails);
    
    // Update student count in the header
    sheet.getRange('B6:C6').setValue(students.length);
    
    // Determine if this is a private lesson
    if (classDetails.isPrivateLesson) {
      populatePrivateLessonData(sheet, classDetails, students);
    } else {
      populateGroupLessonData(sheet, classDetails, students);
    }
  } catch (error) {
    Logger.log(`Error populating template with student data: ${error.message}`);
    throw error;
  }
}

/**
 * Populates the template for a group lesson
 * @param {Sheet} sheet - The sheet to update
 * @param {Object} classDetails - The parsed class details
 * @param {Array} students - Array of student objects
 */
function populateGroupLessonData(sheet, classDetails, students) {
  try {
    // Start at column G (7) for student columns
    let currentColumn = 7;
    
    // For each student, create a pair of B/E columns
    for (let i = 0; i < students.length; i++) {
      // Create merged header for student name
      sheet.getRange(8, currentColumn, 1, 2).merge()
        .setValue(`${students[i].firstName} ${students[i].lastName}`)
        .setFontWeight('bold')
        .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
        .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR)
        .setHorizontalAlignment('center');
      
      // Label the B and E subcolumns
      sheet.getRange(9, currentColumn).setValue('B')
        .setFontWeight('bold')
        .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.BEFORE_COLUMN_COLOR)
        .setHorizontalAlignment('center');
      
      sheet.getRange(9, currentColumn + 1).setValue('E')
        .setFontWeight('bold')
        .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.AFTER_COLUMN_COLOR)
        .setHorizontalAlignment('center');
      
      // Set column widths
      sheet.setColumnWidth(currentColumn, 60);
      sheet.setColumnWidth(currentColumn + 1, 60);
      
      // Move to next student columns
      currentColumn += 2;
    }
    
    // Add first date row at row 10
    const today = new Date();
    sheet.getRange('A10').setValue(today);
    
    // Add 7 more date rows with formula =A10+7, =A11+7, etc.
    for (let i = 1; i < 8; i++) {
      const currentRow = 10 + i;
      const previousRow = currentRow - 1;
      sheet.getRange(`A${currentRow}`).setFormula(`=A${previousRow}+7`);
    }
    
    // Add assessment criteria sections starting at row 18
    // Extract stage from class name
    const stageInfo = extractStageFromClassName(classDetails.program);
    const stageDisplay = stageInfo.value ? `${stageInfo.prefix}${stageInfo.value.toUpperCase()}` : 'S1'; // Default to S1 if no stage found
    
    // Add S1 Pass criteria (rows 18-20)
    sheet.getRange('A18').setValue('Assessment:').setFontWeight('bold');
    sheet.getRange('B18:F18').merge()
      .setValue(`${stageDisplay} Pass`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR);
    
    sheet.getRange('A19').setValue('Criteria:').setFontWeight('bold');
    sheet.getRange('B19:F19').merge()
      .setValue(`${stageDisplay} Repeat`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR);
    
    sheet.getRange('A20').setValue('Notes:').setFontWeight('bold');
    sheet.getRange('B20:F20').merge()
      .setValue(`${stageDisplay} Mid Notes`)
      .setFontWeight('bold')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.SECTION_COLOR);
    
    // Format remaining cells in the assessment rows for data entry
    for (let row = 18; row <= 20; row++) {
      for (let col = 7; col < currentColumn; col++) {
        sheet.getRange(row, col)
          .setHorizontalAlignment('center')
          .setVerticalAlignment('middle');
      }
    }
  } catch (error) {
    Logger.log(`Error populating group lesson data: ${error.message}`);
    throw error;
  }
}

/**
 * Populates the template for a private lesson
 * @param {Sheet} sheet - The sheet to update
 * @param {Object} classDetails - The parsed class details
 * @param {Array} studentsWithDates - Array of student objects with dates
 */
function populatePrivateLessonData(sheet, classDetails, studentsWithDates) {
  try {
    // For private lessons, use a different layout
    // Get students with dates information
    const students = getPrivateLessonStudentsWithDates(classDetails);
    
    // Set up different headers
    sheet.getRange('A8').setValue('Student').setFontWeight('bold');
    sheet.getRange('B8').setValue('Date').setFontWeight('bold');
    sheet.getRange('C8').setValue('Instructor').setFontWeight('bold');
    sheet.getRange('D8').setValue('Notes').setFontWeight('bold');
    sheet.getRange('E8:Z8').setValue('Progress Assessment').setFontWeight('bold');
    
    // Format headers
    sheet.getRange('A8:Z8')
      .setBackground(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_COLOR)
      .setFontColor(DYNAMIC_INSTRUCTOR_CONFIG.CELL_STYLES.HEADER_TEXT_COLOR)
      .setHorizontalAlignment('center');
    
    // Add student data
    for (let i = 0; i < students.length; i++) {
      const rowIndex = 9 + i;
      
      // Student name
      sheet.getRange(rowIndex, 1).setValue(`${students[i].firstName} ${students[i].lastName}`);
      
      // Session date
      if (students[i].date) {
        sheet.getRange(rowIndex, 2).setValue(students[i].date);
      }
      
      // Leave instructor blank for manual entry
      
      // Add alternating row colors
      if (i % 2 === 1) {
        sheet.getRange(rowIndex, 1, 1, 6).setBackground('#f3f3f3');
      }
    }
    
    // Set column widths for private lessons
    sheet.setColumnWidth(1, 150); // Student
    sheet.setColumnWidth(2, 100); // Date
    sheet.setColumnWidth(3, 150); // Instructor
    sheet.setColumnWidth(4, 300); // Notes
    sheet.setColumnWidth(5, 300); // Progress Assessment
  } catch (error) {
    Logger.log(`Error populating private lesson data: ${error.message}`);
    throw error;
  }
}

// Make functions available to other modules
const DynamicInstructorSheet = {
  createDynamicInstructorSheet: createDynamicInstructorSheet,
  onEditDynamicInstructorSheet: onEditDynamicInstructorSheet,
  setupPrivateLessonLayout: setupPrivateLessonLayout,
  rebuildDynamicInstructorSheet: rebuildDynamicInstructorSheet
};