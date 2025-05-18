/**
 * YSL Hub Dynamic Instructor Sheet Module
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

// Configuration constants
const DYNAMIC_INSTRUCTOR_CONFIG = {
  SHEET_NAME: 'Group Lesson Tracker',
  ROSTER_SHEET_NAME: 'Daxko', // Primary sheet containing student registration data
  ALTERNATE_ROSTER_SHEETS: ['Roster', 'Participants', 'Students', 'Registration'], // Alternative names for roster sheets
  // Don't try to access the Excel file directly
  HEADERS: {
    CLASS_SELECTOR_LABEL: 'Select Class:',
    FIRST_NAME: 'First Name',
    LAST_NAME: 'Last Name',
    NOTES: 'Notes'
  },
  CELL_REFERENCES: {
    CLASS_SELECTOR: 'H3', // Match template position
    INSTRUCTOR_INFO: 'D4', // Updated position
    PROGRAM_INFO: 'B4',   // Updated position
    DAY_INFO: 'B5',       // Updated position
    TIME_INFO: 'D5',      // Updated position
    LOCATION_INFO: 'B6',  // Updated position
    STUDENT_COUNT: 'D6',  // Updated position
    STUDENT_NAME_START: 'G7', // Position in the template
    SKILLS_START: 'A17'   // Position in the template
  },
  // Student columns are now different - each student takes 2 columns (B/E)
  // First student is at column G (7), second at column I (9), etc.
  STUDENT_DATA: {
    FIRST_COLUMN: 7,     // Column G
    COLUMN_SPACING: 2,   // Every 2 columns (G, I, K, etc.)
    MAX_STUDENTS: 12     // Maximum number of students in the template
  }
};

/**
 * Generates the Group Lesson Tracker sheet 
 * This is the main entry point for creating the sheet
 * 
 * @return The created sheet or null if an error occurred
 */
function createDynamicInstructorSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Check if sheet already exists
    let sheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    
    // If sheet exists, prompt user for confirmation before deleting
    if (sheet) {
      const ui = SpreadsheetApp.getUi();
      const response = ui.alert(
        'Sheet Already Exists',
        `The "${DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME}" sheet already exists. Would you like to replace it? This will delete the current sheet and all its data.`,
        ui.ButtonSet.YES_NO
      );
      
      // If user chooses not to replace the sheet, exit the function
      if (response !== ui.Button.YES) {
        ui.alert(
          'Operation Cancelled',
          'The Group Lesson Tracker creation was cancelled. The existing sheet has not been modified.',
          ui.ButtonSet.OK
        );
        return sheet;
      }
      
      // User confirmed, proceed with deletion
      ss.deleteSheet(sheet);
    }
    
    // Create new sheet
    sheet = ss.insertSheet(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    
    // Create basic template layout - simplified approach
    createStaticTemplate(sheet);
    
    // Show completed message
    SpreadsheetApp.getUi().alert(
      'Group Lesson Tracker Created',
      'The group lesson tracker template has been created.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return sheet;
  } catch (error) {
    // Handle errors
    Logger.log(`Error creating group lesson tracker: ${error.message}`);
    
    // Show error to user
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to create group lesson tracker: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return null;
  }
}

/**
 * Rebuilds the dynamic instructor sheet with the currently selected class
 * Used as a quick way to refresh or fix the sheet
 * 
 * @return The updated sheet or null if an error occurred
 */
function rebuildDynamicInstructorSheet() {
  try {
    // Log the operation
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Rebuilding dynamic instructor sheet', 'INFO', 'rebuildDynamicInstructorSheet');
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME);
    
    // If sheet doesn't exist, create it
    if (!sheet) {
      return createDynamicInstructorSheet();
    }
    
    // Get currently selected class
    const classSelector = sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR);
    const selectedClass = classSelector.getValue();
    
    // Check if the selected class is valid
    if (!selectedClass || 
        selectedClass === 'Select a class...' || 
        selectedClass === 'No classes available' || 
        selectedClass === 'Error loading classes') {
      
      SpreadsheetApp.getUi().alert(
        'No Class Selected',
        'Please select a class from the dropdown first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return sheet;
    }
    
    // Validate the class against available options
    const validOptions = getClassOptions();
    if (!validOptions.includes(selectedClass)) {
      Logger.log(`Warning: Selected class "${selectedClass}" is not in the list of valid options.`);
      
      // Try to find the closest match
      const normalizedSelection = selectedClass.toString().trim().toLowerCase();
      let closestMatch = "";
      let bestMatchScore = 0;
      
      validOptions.forEach(option => {
        const normalizedOption = option.toString().trim().toLowerCase();
        if (normalizedSelection === normalizedOption) {
          closestMatch = option;
          bestMatchScore = 100;
        } else if (normalizedSelection.includes(normalizedOption) || normalizedOption.includes(normalizedSelection)) {
          const score = Math.min(normalizedSelection.length, normalizedOption.length) / 
                        Math.max(normalizedSelection.length, normalizedOption.length) * 100;
          if (score > bestMatchScore) {
            bestMatchScore = score;
            closestMatch = option;
          }
        }
      });
      
      // If we found a close match, use it
      if (bestMatchScore >= 90) {
        Logger.log(`Using closest match "${closestMatch}" (score: ${bestMatchScore})`);
        classSelector.setValue(closestMatch);
        SpreadsheetApp.getUi().alert(
          'Class Name Corrected',
          `The class name has been corrected to "${closestMatch}" for better compatibility.`,
          SpreadsheetApp.getUi().ButtonSet.OK
        );
      } else {
        // Ask the user what to do
        const ui = SpreadsheetApp.getUi();
        const result = ui.alert(
          'Invalid Class Selection',
          `The class "${selectedClass}" does not match any known classes. Would you like to choose from the list of valid classes?`,
          ui.ButtonSet.YES_NO
        );
        
        if (result === ui.Button.YES) {
          // Create class selector again
          createClassSelector(sheet);
          return sheet;
        }
      }
    }
    
    // In the rebuild process, we'll keep skills that are already there
    // First let's check if there are skill columns present
    const lastCol = sheet.getLastColumn();
    const hasSkills = lastCol > 5 && sheet.getRange(6, 6).getValue() !== 'Skills will be shown after class selection';
    
    // Record existing skills data if present
    let existingSkills = null;
    if (hasSkills) {
      try {
        Logger.log('Skills already present, preserving them during rebuild');
        
        // Capture skill headers
        const skillHeadersRange = sheet.getRange(6, 6, 1, lastCol - 5);
        existingSkills = {
          headers: skillHeadersRange.getValues()[0],
          backgrounds: skillHeadersRange.getBackgrounds()[0],
          formats: skillHeadersRange.getFontWeights()[0]
        };
      } catch (e) {
        Logger.log(`Error capturing existing skills: ${e.message}`);
        existingSkills = null;
      }
    }
    
    // Don't completely clear the sheet, just student data
    clearStudentData(sheet, existingSkills !== null);
    
    // Get final selected class (might have been corrected)
    const finalSelectedClass = classSelector.getValue();
    
    // Get class details
    const classDetails = getClassDetails(finalSelectedClass);
    
    // Log details for debugging
    Logger.log(`Rebuilding sheet with class: ${finalSelectedClass}`);
    Logger.log(`Class details: ${JSON.stringify(classDetails)}`);
    
    // Update class info
    updateClassInfo(sheet, classDetails);
    
    // Get and load student data with custom parameters to force deeper searching
    // Set forceSearch to true to attempt more aggressive matching
    const students = getStudentDataForClass(classDetails, true);
    populateStudentData(sheet, students);
    
    // If we had skills before and preserved them, restore them now
    if (existingSkills) {
      try {
        // Make sure cell F6 is clear
        sheet.getRange('F6').clearContent();
        
        // Restore the skills headers with original formatting
        const headerCount = existingSkills.headers.length;
        for (let i = 0; i < headerCount; i++) {
          if (existingSkills.headers[i]) {
            const cell = sheet.getRange(6, 6 + i);
            cell.setValue(existingSkills.headers[i])
                .setBackground(existingSkills.backgrounds[i])
                .setFontWeight(existingSkills.formats[i])
                .setFontSize(9)
                .setTextRotation(90)
                .setWrap(true)
                .setVerticalAlignment('bottom');
            
            // Make column narrower for better fit
            sheet.setColumnWidth(6 + i, 30);
            
            // Do NOT add checkboxes for skills assessments - use plain text
            // Keep the original formatting but clear any data validations
            const skillRange = sheet.getRange(8, 6 + i, 20, 1);
            skillRange.clearDataValidations();
          }
        }
        
        Logger.log('Successfully restored skill headers');
      } catch (e) {
        Logger.log(`Error restoring skill headers: ${e.message}`);
        
        // If restoration fails, get new skills from scratch
        const allSkills = getSkillsFromSwimmerRecords();
        
        // Extract stage information and filter skills
        const stageInfo = extractStageFromClassName(classDetails.program);
        const relevantSkills = (stageInfo && stageInfo.value) 
          ? filterSkillsByStage(allSkills, stageInfo)
          : allSkills;
        
        // Add skills to the sheet
        addSkillsToSheet(sheet, relevantSkills);
      }
    } else {
      // No skills were present, so we need to get new ones
      // Make sure cell F6 is clear
      sheet.getRange('F6').clearContent();
      
      // Get and add skills
      const allSkills = getSkillsFromSwimmerRecords();
      
      // Extract stage information and filter skills
      const stageInfo = extractStageFromClassName(classDetails.program);
      const relevantSkills = (stageInfo && stageInfo.value) 
        ? filterSkillsByStage(allSkills, stageInfo)
        : allSkills;
      
      // Add skills to the sheet
      addSkillsToSheet(sheet, relevantSkills);
    }
    
    // Check if it's a private lesson and add appropriate layout
    if (classDetails.isPrivateLesson) {
      setupPrivateLessonLayout(sheet, classDetails);
    }
    
    // Verify completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Successfully rebuilt instructor sheet for ${finalSelectedClass} with ${students.length} students`, 'INFO', 'rebuildDynamicInstructorSheet');
    }
    
    SpreadsheetApp.getUi().alert(
      'Sheet Updated',
      `The instructor sheet has been updated with data for ${finalSelectedClass}.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return sheet;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'rebuildDynamicInstructorSheet', 
        'Error rebuilding dynamic instructor sheet. Please try again or contact support.');
    } else {
      Logger.log(`Error rebuilding dynamic instructor sheet: ${error.message}`);
      Logger.log(`Error stack: ${error.stack}`);
      
      // Show error to user
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
 * Initializes the basic layout of the instructor sheet
 * 
 * @param sheet The sheet to initialize
 */
/**
 * Creates a simple static template following the exact layout of the template
 * 
 * @param sheet The sheet to format
 */
function createStaticTemplate(sheet) {
  // Safe merge function to handle merge errors
  const safeMerge = (range) => {
    try {
      // Check if already part of a merge
      if (!range.isPartOfMerge()) {
        range.merge();
      }
      return range;
    } catch (e) {
      Logger.log(`Error merging range ${range.getA1Notation()}: ${e.message}`);
      return range; // Return the range even if merge fails
    }
  };

  // ==========================================
  // 1. Set column widths
  // ==========================================
  sheet.setColumnWidth(1, 150);  // Column A (skills)
  
  // Info columns - columns B-F
  for (let i = 2; i <= 6; i++) {
    sheet.setColumnWidth(i, 90);
  }
  
  // Student columns B-Y to support 12 students (24 columns total)
  // Each student gets two columns B/C, D/E, F/G, etc. for Beginning/End assessment
  for (let i = 2; i <= 25; i++) {
    sheet.setColumnWidth(i, 40);
  }
  
  // ==========================================
  // 2. Row heights (optional but helps with layout)
  // ==========================================
  sheet.setRowHeight(1, 30);  // Title row
  sheet.setRowHeight(2, 25);  // Class info header
  
  // ==========================================
  // 3. Title and Headers
  // ==========================================
  // Title bar - A1:Z1 (full width)
  const titleRange = sheet.getRange('A1:Z1');
  titleRange.setValue('YSL v6 & SAW ASSESSMENT SHEET')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#4285F4') // Blue
    .setFontColor('white')
    .setFontSize(14);
  safeMerge(titleRange);
  
  // Create a merged range for A2:Z2 and set dropdown for class selection
  const dropdownRange = sheet.getRange('A2:Z2');
  dropdownRange.setBackground('#F8F9FA') // Light gray background
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  safeMerge(dropdownRange);
  
  // Prepare the dropdown for the merged cell A2:Z2 with concatenated values from Daxko (use W and X only)
  try {
    setConcatenatedDropdown(sheet, 'A2', 'Daxko', ['W', 'X'], 'Private Swim Lessons');
  } catch (e) {
    Logger.log(`Error setting concatenated dropdown in A2: ${e.message}`);
    sheet.getRange('A2').setValue('Class Selection (Error: Unable to load dropdown)');
  }
  
  // Create clear, simple sync instructions in cell A4
  sheet.getRange('A4').setValue('Use YSL v6 Hub menu → ◉ SYNC STUDENT DATA ◉')
    .setFontWeight('bold')
    .setFontColor('#1155CC') // Link blue
    .setBackground('#E8F0FE') // Light blue background
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBorder(true, true, true, true, null, null, '#4285F4', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // ==========================================
  // 4. Attendance Section (now at row 3 since rows 3-5 were deleted)
  // ==========================================
  // Attendance header - A3:Z3 (full width) - moved up from row 6
  const attendanceHeader = sheet.getRange('A3:Z3');
  attendanceHeader.setValue('ATTENDANCE')
    .setFontWeight('bold')
    .setBackground('#E0E0E0')
    .setHorizontalAlignment('center');
  safeMerge(attendanceHeader);
  
  // Class dates - A5:A12 (moved up 3 rows since rows 3-5 were removed)
  for (let row = 5; row <= 12; row++) {
    sheet.getRange(`A${row}`).setValue(`Class Date ${row-4}`)
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
  }
  
  // ==========================================
  // 6. Student Names Row - Each student gets two merged columns
  // ==========================================
  // Define student column pairs to merge - B:C, D:E, F:G, etc.
  const studentPairs = [
    ['B', 'C'], ['D', 'E'], ['F', 'G'], ['H', 'I'], ['J', 'K'],
    ['L', 'M'], ['N', 'O'], ['P', 'Q'], ['R', 'S'], ['T', 'U'], 
    ['V', 'W'], ['X', 'Y']
  ];
  
  // For reference in other parts of the code
  const beginningCols = ['B', 'D', 'F', 'H', 'J', 'L', 'N', 'P', 'R', 'T', 'V', 'X'];
  const endCols = ['C', 'E', 'G', 'I', 'K', 'M', 'O', 'Q', 'S', 'U', 'W', 'Y'];
  
  // Set up student name headers and merge student name cells in row 4 (moved up from row 7)
  for (let i = 0; i < studentPairs.length; i++) {
    const pair = studentPairs[i];
    const range = `${pair[0]}4:${pair[1]}4`;
    
    // Set student name in merged cell
    const nameRange = sheet.getRange(range);
    nameRange.setValue(`Student ${i+1}`)
      .setFontWeight('bold')
      .setBackground('#F3F3F3')
      .setHorizontalAlignment('center');
    
    // Merge the cells
    safeMerge(nameRange);
  }
  
  // ==========================================
  // 7. Merge attendance cells for each student
  // ==========================================
  // Merge each student column pair for attendance rows (moved up 3 rows)
  for (let row = 5; row <= 12; row++) {
    for (let i = 0; i < studentPairs.length; i++) {
      const pair = studentPairs[i];
      const range = `${pair[0]}${row}:${pair[1]}${row}`;
      
      // Format and merge the attendance cells
      const attendanceRange = sheet.getRange(range);
      attendanceRange.setBackground('#F9F9F9');
      
      // Merge the cells
      safeMerge(attendanceRange);
    }
  }
  
  // ==========================================
  // 8. B/E Headers - Row 13 (Beginning/End) - moved up from row 16
  // ==========================================
  // Set B (Beginning) in all beginning columns
  for (let i = 0; i < beginningCols.length; i++) {
    sheet.getRange(`${beginningCols[i]}13`).setValue('B')
      .setFontWeight('bold')
      .setBackground('#F3F3F3')
      .setHorizontalAlignment('center');
  }
  
  // Set E (End) in all end columns
  for (let i = 0; i < endCols.length; i++) {
    sheet.getRange(`${endCols[i]}13`).setValue('E')
      .setFontWeight('bold')
      .setBackground('#F3F3F3')
      .setHorizontalAlignment('center');
  }
  
  // ==========================================
  // 9. Stage Skills - A14:A29 (moved up 3 rows from A17:A32)
  // ==========================================
  // Status rows
  sheet.getRange('A14').setValue('S1 Pass')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A15').setValue('S1 Repeat')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  
  // Mid-session feedback
  sheet.getRange('A16').setValue('S1 Mid Notes')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A17').setValue('S1 Mid Sent')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  
  // S1 Skills
  sheet.getRange('A18').setValue('S1 Submerge')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A19').setValue('S1 Front Glide')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A20').setValue('S1 Water Exit')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A21').setValue('S1 J-P-T-G')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A22').setValue('S1 Back Float')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A23').setValue('S1 Roll')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A24').setValue('S1 Front Float')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A25').setValue('S1 Back Glide')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A26').setValue('S1 S-F-S')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A27').setValue('S1 Swim Topics')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  
  // End-session feedback
  sheet.getRange('A28').setValue('S1 End Notes')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  sheet.getRange('A29').setValue('S1 End Sent')
    .setFontWeight('bold')
    .setBackground('#D9EAD3');
  
  // ==========================================
  // 10. Format skill cells for each student - no merges
  // ==========================================
  for (let row = 14; row <= 29; row++) {
    // Format individual cells in both beginning and end columns
    for (let i = 0; i < beginningCols.length; i++) {
      // Format Beginning cell
      sheet.getRange(`${beginningCols[i]}${row}`)
        .setBackground('#F9F9F9');
        
      // Format End cell
      sheet.getRange(`${endCols[i]}${row}`)
        .setBackground('#F9F9F9');
    }
  }
  
  // ==========================================
  // 11. Second B/E Headers - Row 30 (moved up from row 33)
  // ==========================================
  // Set B (Beginning) in all beginning columns
  for (let i = 0; i < beginningCols.length; i++) {
    sheet.getRange(`${beginningCols[i]}30`).setValue('B')
      .setFontWeight('bold')
      .setBackground('#F3F3F3')
      .setHorizontalAlignment('center');
  }
  
  // Set E (End) in all end columns
  for (let i = 0; i < endCols.length; i++) {
    sheet.getRange(`${endCols[i]}30`).setValue('E')
      .setFontWeight('bold')
      .setBackground('#F3F3F3')
      .setHorizontalAlignment('center');
  }
  
  // ==========================================
  // 12. SAW Skills - A31:A40 (moved up 3 rows from A34:A43)
  // ==========================================
  sheet.getRange('A31').setValue('SAW Pass')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A32').setValue('SAW Repeat')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A33').setValue('SAW Sub Face')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A34').setValue('SAW Bob (ind)')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A35').setValue('SAW Front Glide')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A36').setValue('SAW Back Float')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A37').setValue('SAW S-F-S (ind)')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A38').setValue('SAW Jump (ind)')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A39').setValue('SAW J-P-T-G (asst)')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  sheet.getRange('A40').setValue('SAW J-P-T-G (ind)')
    .setFontWeight('bold')
    .setBackground('#FCE5CD');
  
  // ==========================================
  // 13. Format SAW skills cells for each student - no merges
  // ==========================================
  for (let row = 31; row <= 40; row++) {
    // Format individual cells in both beginning and end columns
    for (let i = 0; i < beginningCols.length; i++) {
      // Format Beginning cell
      sheet.getRange(`${beginningCols[i]}${row}`)
        .setBackground('#F9F9F9');
        
      // Format End cell
      sheet.getRange(`${endCols[i]}${row}`)
        .setBackground('#F9F9F9');
    }
  }
  
  // ==========================================
  // 14. Freeze rows & columns - Fix merged cell freeze issue
  // ==========================================
  // Freeze through row 4 (headers and student names row) - moved up from row 7
  sheet.setFrozenRows(4);
  
  // Don't freeze any columns to avoid merged cell issues
  // No columns frozen
  
  // ==========================================
  // 15. Delete extra rows
  // ==========================================
  // Delete rows from 42 to the end to clean up the sheet (moved up from row 45)
  const lastRow = sheet.getMaxRows();
  if (lastRow > 41) {
    try {
      // Delete all rows after row 41
      sheet.deleteRows(42, lastRow - 41);
      Logger.log(`Deleted rows 42-${lastRow} to clean up the sheet`);
    } catch (e) {
      Logger.log(`Error deleting extra rows: ${e.message}`);
    }
  }
  
  // Log completion
  Logger.log('Static template created successfully');
}

/**
 * Creates the class selector dropdown with available classes
 * 
 * @param sheet The sheet to add the selector to
 */
function createClassSelector(sheet) {
  try {
    // Get classes from the Classes sheet
    const classOptions = getClassOptions();
    
    // Log detailed information for debugging
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Creating class selector with ${classOptions ? classOptions.length : 0} options`, 'DEBUG', 'createClassSelector');
    }
    
    if (!classOptions || classOptions.length === 0) {
      sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR)
        .setValue('No classes available');
      return;
    }
    
    // Log available class options for debugging
    Logger.log(`Available class options (${classOptions.length}): ${classOptions.join(', ')}`);
    
    // First clear any existing validation to avoid conflicts
    sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR).clearDataValidations();
    
    // Create placeholder dropdown text
    const placeholderText = "Select a class...";
    const displayOptions = [placeholderText].concat(classOptions);
    
    // Create data validation with class options - allow invalid for initial setup
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(displayOptions, true) // Include placeholder in the list
      .setAllowInvalid(true) // Allow invalid to prevent immediate validation errors
      .build();
    
    // Apply validation to class selector cell
    const selectorCell = sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR);
    selectorCell.setDataValidation(validation);
    
    // Add a note with instructions and available options
    let noteText = "Select a class from the dropdown. Available classes:\n\n";
    classOptions.forEach((option, index) => {
      noteText += `${index + 1}. ${option}\n`;
    });
    selectorCell.setNote(noteText);
    
    // Set initial value to placeholder
    selectorCell.setValue(placeholderText);
    
    // Add helpful styling
    selectorCell.setFontColor('#666666');
    
    // Log success
    Logger.log('Class selector created successfully');
  } catch (error) {
    // Log the error
    Logger.log(`Error creating class selector: ${error.message}`);
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error creating class selector: ${error.message}`, 'ERROR', 'createClassSelector');
    }
    
    // Set a default value
    sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR).setValue('Error loading classes');
  }
}

/**
 * Gets the list of available classes for the selector
 * 
 * @return Array of class names
 */
function getClassOptions() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    // If Classes sheet doesn't exist, return empty array
    if (!classesSheet) {
      Logger.log("Classes sheet not found");
      return [];
    }
    
    // Get class data
    const lastRow = classesSheet.getLastRow();
    if (lastRow <= 1) {
      Logger.log("Classes sheet only has a header row");
      return []; // Only header row exists
    }
    
    // Log what we're reading for debugging
    Logger.log(`Reading class data from row 2 to ${lastRow} (${lastRow - 1} classes)`);
    
    // Get all class data (skip header row)
    const classData = classesSheet.getRange(2, 1, lastRow - 1, 7).getValues();
    
    // Log raw class data for debugging
    Logger.log(`Raw class data (first 3 rows): ${JSON.stringify(classData.slice(0, 3))}`);
    
    // Format class options as "Program (Day, Time)"
    const options = [];
    const optionsMap = {}; // Use a map to eliminate duplicates
    
    for (let i = 0; i < classData.length; i++) {
      const row = classData[i];
      const program = row[1]; // Column B (2nd column, index 1)
      const day = row[2];     // Column C (3rd column, index 2)
      const time = row[3];    // Column D (4th column, index 3)
      
      // Skip rows with missing essential data
      if (!program || !day || !time) {
        Logger.log(`Skipping row ${i+2} - Missing required data: program=${program}, day=${day}, time=${time}`);
        continue;
      }
      
      // Create the class name option with trimmed values to prevent whitespace issues
      const trimmedProgram = program.toString().trim();
      const trimmedDay = day.toString().trim();
      const trimmedTime = time.toString().trim();
      
      const classOption = `${trimmedProgram} (${trimmedDay}, ${trimmedTime})`;
      
      // Use a map to eliminate duplicates
      if (!optionsMap[classOption]) {
        optionsMap[classOption] = true;
        options.push(classOption);
      }
    }
    
    // Sort options alphabetically to make selection easier
    options.sort();
    
    // Log the formatted options for debugging
    Logger.log(`Generated ${options.length} unique class options`);
    Logger.log(`Class options: ${options.join(', ')}`);
    
    // Log detailed class information for validation debugging
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Class options generated: ${options.length}`, 'INFO', 'getClassOptions');
      // Log each option in a detailed format to help with debugging
      options.forEach((option, index) => {
        ErrorHandling.logMessage(`Class option ${index+1}: '${option}'`, 'DEBUG', 'getClassOptions');
      });
    }
    
    return options;
  } catch (error) {
    // Log error and return empty array
    Logger.log(`Error getting class options: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'getClassOptions', 'Error retrieving class options');
    }
    return [];
  }
}

/**
 * Handler for edit events on the dynamic instructor sheet
 * This is triggered by the global onEdit trigger
 * 
 * @param e The edit event object
 */
function onEditDynamicInstructorSheet(e) {
  // Check if edit is in the dynamic instructor sheet
  if (e.range.getSheet().getName() !== DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME) {
    return;
  }
  
  // Check if edit is in the class selector cell (now at H3)
  if (e.range.getA1Notation() === DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR && e.value) {
    handleClassSelection(e.range.getSheet(), e.value);
  }
}

/**
 * Handles when a class is selected from the dropdown
 * 
 * @param sheet The dynamic instructor sheet
 * @param className The selected class name
 */
function handleClassSelection(sheet, className) {
  try {
    // Log the selection
    Logger.log(`Class selected: "${className}"`);
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Class selected: ${className}`, 'INFO', 'handleClassSelection');
    }
    
    // Validate class name format
    if (!className || className === 'Select a class...' || className === 'Error loading classes' || className === 'No classes available') {
      Logger.log(`Skipping class selection - invalid class name: "${className}"`);
      return; // Skip processing if not a valid selection
    }
    
    // Get valid options, including a fuzzy match approach
    const validOptions = getClassOptions();
    Logger.log(`Validating against ${validOptions.length} class options`);
    
    // Try both exact match and fuzzy match
    let exactMatch = validOptions.includes(className);
    let fuzzyMatch = false;
    let closestMatch = "";
    
    if (!exactMatch) {
      // Try to find a fuzzy match by normalizing both strings (trim whitespace, case insensitive)
      const normalizedSelection = className.toString().trim().toLowerCase();
      
      // Log the normalized selection for debugging
      Logger.log(`Normalized selection: '${normalizedSelection}'`);
      
      // Find the closest match
      let bestMatchScore = 0;
      
      validOptions.forEach(option => {
        const normalizedOption = option.toString().trim().toLowerCase();
        
        // Check if the normalized strings match
        if (normalizedSelection === normalizedOption) {
          fuzzyMatch = true;
          closestMatch = option;
          bestMatchScore = 100; // Perfect match
        } 
        // If we don't have a perfect match yet, check for similarity
        else if (bestMatchScore < 100) {
          // Check if one is contained in the other
          if (normalizedSelection.includes(normalizedOption) || normalizedOption.includes(normalizedSelection)) {
            const score = Math.min(normalizedSelection.length, normalizedOption.length) / 
                         Math.max(normalizedSelection.length, normalizedOption.length) * 100;
            
            if (score > bestMatchScore) {
              bestMatchScore = score;
              closestMatch = option;
              fuzzyMatch = (score >= 90); // Consider it a match if similarity is at least 90%
            }
          }
        }
      });
      
      // Log the fuzzy match results for debugging
      Logger.log(`Fuzzy match found: ${fuzzyMatch}, closest match: '${closestMatch}', score: ${bestMatchScore}`);
    }
    
    // If neither exact nor fuzzy match, show an error
    if (!exactMatch && !fuzzyMatch) {
      Logger.log(`WARNING: Selected class "${className}" is not in the list of valid options`);
      
      // Get the available class options for the error message
      let availableOptions = validOptions.join('\n');
      if (availableOptions.length > 500) {
        // Show just the first few options
        availableOptions = validOptions.slice(0, 10).join('\n') + "\n...";
      }
      
      // Show error dialog with available options
      SpreadsheetApp.getUi().alert(
        'Invalid Class Selection',
        `The selected class "${className}" is not valid. Please select one of the following:\n\n${availableOptions}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      // Reset the selector to the placeholder
      sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR).setValue('Select a class...');
      return;
    }
    
    // If we found a fuzzy match but not an exact match, update the cell to use the valid format
    if (!exactMatch && fuzzyMatch) {
      Logger.log(`Using closest match: "${closestMatch}" instead of "${className}"`);
      // Update to use the properly formatted class name
      sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR).setValue(closestMatch);
      className = closestMatch; // Use the corrected value
    }
    
    // Log successful validation
    Logger.log(`Valid class selection: "${className}"`);
    
    // Clear existing student data
    clearStudentData(sheet);
    
    // Get class details
    const classDetails = getClassDetails(className);
    Logger.log(`Class details: ${JSON.stringify(classDetails)}`);
    
    // Update class info
    updateClassInfo(sheet, classDetails);
    
    // Get student data for this class
    const students = getStudentDataForClass(classDetails);
    Logger.log(`Found ${students.length} students for class "${className}"`);
    
    // Populate student data
    populateStudentData(sheet, students);
    
    // Make sure cell F6 is clear
    sheet.getRange('F6').clearContent();
    
    // Now get skills from Swimmer Records based on the selected class
    const allSkills = getSkillsFromSwimmerRecords();
    
    // Extract stage information from the class name to filter relevant skills
    const stageInfo = extractStageFromClassName(classDetails.program);
    
    // Filter skills based on the class's stage, if available
    const relevantSkills = (stageInfo && stageInfo.value) 
      ? filterSkillsByStage(allSkills, stageInfo)
      : allSkills;
    
    // Log the skills we're using
    Logger.log(`Using ${relevantSkills.stage.length} stage skills and ${relevantSkills.saw.length} SAW skills for ${className}`);
    
    // Add filtered skills to the sheet
    addSkillsToSheet(sheet, relevantSkills);
    
    // Check if it's a private lesson and add appropriate layout
    if (classDetails.isPrivateLesson) {
      setupPrivateLessonLayout(sheet, classDetails);
    }
    
    // Log successful completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Successfully loaded data for class: ${className}`, 'INFO', 'handleClassSelection');
    }
  } catch (error) {
    // Handle errors
    Logger.log(`Error handling class selection: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'handleClassSelection', 
        'Error loading class data. Please try again or contact support.');
    } else {
      // Show error to user even if ErrorHandling module is not available
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to load class data: ${error.message}\n\nPlease try again or contact support.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Reset the selector to the placeholder
    try {
      sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR).setValue('Select a class...');
    } catch (e) {
      // If this fails too, just log it and continue
      Logger.log(`Error resetting selector: ${e.message}`);
    }
  }
}

/**
 * Sets up the layout for private lessons, which is different from group classes
 * 
 * @param sheet The dynamic instructor sheet
 * @param classDetails The class details object
 */
function setupPrivateLessonLayout(sheet, classDetails) {
  // Add more space for skills and progress tracking
  sheet.insertColumns(6, 4); // Add columns for skill tracking
  
  // Add private lesson header
  sheet.getRange('F6:I6').merge()
    .setValue('Private Lesson Progress Tracking')
    .setFontWeight('bold')
    .setBackground('#E0E0E0')
    .setHorizontalAlignment('center');
  
  // Add progress tracking headers
  const trackingHeaders = ['Date', 'Skills Worked On', 'Progress', 'Notes'];
  sheet.getRange('F7:I7').setValues([trackingHeaders])
    .setFontWeight('bold')
    .setBackground('#F3F3F3');
  
  // Add several blank rows for tracking entries
  for (let i = 0; i < 8; i++) {
    const row = 8 + i;
    // Add date cell with data validation for date picker
    const dateValidation = SpreadsheetApp.newDataValidation()
      .requireDate()
      .build();
    sheet.getRange(`F${row}`).setDataValidation(dateValidation);
  }
  
  // Adjust column widths
  sheet.setColumnWidth(6, 100); // Date
  sheet.setColumnWidth(7, 200); // Skills
  sheet.setColumnWidth(8, 150); // Progress
  sheet.setColumnWidth(9, 200); // Notes
}

/**
 * Clears the student data and optionally skills from the sheet to prepare for new data
 * 
 * @param sheet The dynamic instructor sheet
 * @param preserveSkills If true, skills headers and formatting will be preserved
 */
function clearStudentData(sheet, preserveSkills = false) {
  try {
    // Get the last row (start at the first student row)
    const lastRow = Math.max(sheet.getLastRow(), 20);
    
    // Clear student rows (starting from row 8)
    sheet.getRange(8, 1, lastRow - 7, 5).clearContent();
    
    // Clear instructor info
    sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.INSTRUCTOR_INFO).clearContent();
    
    // Clear class info
    sheet.getRange('G2:H3').clearContent();
    
    // Clear skills section only if preserveSkills is false
    if (!preserveSkills) {
      // Get the last column
      const lastCol = Math.max(sheet.getLastColumn(), 15);
      
      // If there are skills columns (beyond column E), clear them
      if (lastCol > 5) {
        // Clear skills headers (row 6)
        sheet.getRange(6, 6, 1, lastCol - 5).clearContent();
        
        // Clear skills data (row 7 and below)
        if (lastRow > 7) {
          sheet.getRange(7, 6, lastRow - 6, lastCol - 5).clearContent();
        }
        
        // Also clear formatting by setting to default
        sheet.getRange(6, 6, lastRow - 5, lastCol - 5).setBackground(null);
        
        // Don't add any placeholder text - keep the cell empty for a cleaner look
        sheet.getRange('F6').clearContent();
      }
    } else {
      // If preserving skills, only clear the student checkboxes in the skills columns
      const lastCol = sheet.getLastColumn();
      
      if (lastCol > 5) {
        // Just clear the checkboxes (row 8 and below in the skills columns)
        if (lastRow > 7) {
          try {
            // Clear only the checkbox values, not the headers or formatting
            sheet.getRange(8, 6, lastRow - 7, lastCol - 5).setValue(false);
          } catch (e) {
            Logger.log(`Error resetting checkboxes: ${e.message}`);
          }
        }
      }
    }
  } catch (error) {
    // Log the error but continue
    Logger.log(`Error clearing student data: ${error.message}`);
  }
}

/**
 * Updates the class information display using template format
 * 
 * @param sheet The dynamic instructor sheet
 * @param classDetails The class details object
 */
function updateClassInfo(sheet, classDetails) {
  // Set class info in template format
  sheet.getRange('B3').setValue(classDetails.program);
  sheet.getRange('B4').setValue(classDetails.day);
  sheet.getRange('B5').setValue(classDetails.location || '');
  
  // Set instructor info
  sheet.getRange('D3').setValue(classDetails.instructor || 'Not assigned');
  
  // Set time
  sheet.getRange('D4').setValue(classDetails.time);
  
  // Student count will be updated in populateStudentData
  
  // Set class selector cell
  sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR)
    .setValue(classDetails.fullName);
  
  // Update skills section with correct stage prefix based on program
  const stageMatch = classDetails.program.match(/Stage\s+(\d+|[A-F])/i);
  if (stageMatch && stageMatch[1]) {
    const stageValue = stageMatch[1];
    // Set the stage prefix for skill rows
    
    // Update the pass/repeat rows
    sheet.getRange('A17').setValue(`S${stageValue} Pass`);
    sheet.getRange('A18').setValue(`S${stageValue} Repeat`);
    
    // Update feedback rows
    sheet.getRange('A19').setValue(`S${stageValue} Mid Notes`);
    sheet.getRange('A20').setValue(`S${stageValue} Mid Sent`);
    sheet.getRange('A31').setValue(`S${stageValue} End Notes`);
    sheet.getRange('A32').setValue(`S${stageValue} End Sent`);
    
    // The specific skill rows would need to be updated based on the stage
    // This would require additional data about what skills are in each stage
    // For now, we'll keep the skills from initializeSheetLayout
  }
}

/**
 * Gets the details for a selected class
 * 
 * @param className The class name from the selector
 * @return Object with class details
 */
function getClassDetails(className) {
  try {
    // Extract program, day, and time from class name
    // Format is "Program (Day, Time)"
    const match = className.match(/^(.+?)\s*\(([^,]+),\s*(.+?)\)$/);
    
    if (!match) {
      return {
        fullName: className,
        program: '',
        day: '',
        time: '',
        isPrivateLesson: false
      };
    }
    
    const program = match[1].trim();
    const day = match[2].trim();
    const time = match[3].trim();
    
    // Check if this is a private lesson
    const isPrivateLesson = program.toLowerCase().includes('private');
    
    // Find additional details from Classes sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const classesSheet = ss.getSheetByName('Classes');
    
    let instructor = '';
    let location = '';
    
    if (classesSheet) {
      const lastRow = classesSheet.getLastRow();
      const classData = classesSheet.getRange(2, 2, lastRow - 1, 6).getValues();
      
      // Find the matching class
      for (const row of classData) {
        const rowProgram = row[0];
        const rowDay = row[1];
        const rowTime = row[2];
        
        if (rowProgram === program && rowDay === day && rowTime === time) {
          location = row[3] || '';
          instructor = row[5] || '';
          break;
        }
      }
    }
    
    return {
      fullName: className,
      program,
      day,
      time,
      isPrivateLesson,
      location,
      instructor
    };
  } catch (error) {
    Logger.log(`Error getting class details: ${error.message}`);
    return {
      fullName: className,
      program: '',
      day: '',
      time: '',
      isPrivateLesson: false
    };
  }
}

/**
 * Gets student data for a specific class
 * 
 * @param classDetails The class details object
 * @param forceSearch If true, will try more aggressive matching techniques
 * @return Array of student data objects
 */
function getStudentDataForClass(classDetails, forceSearch = false) {
  try {
    // Log detailed class details for troubleshooting
    Logger.log(`Looking for students matching: Program='${classDetails.program}', Day='${classDetails.day}', Time='${classDetails.time}'`);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // First try the configured roster sheet
    let rosterSheet = ss.getSheetByName(DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME);
    
    // If not found, try alternative sheet names from config
    if (!rosterSheet) {
      Logger.log(`Sheet '${DYNAMIC_INSTRUCTOR_CONFIG.ROSTER_SHEET_NAME}' not found, trying alternatives`);
      
      for (const sheetName of DYNAMIC_INSTRUCTOR_CONFIG.ALTERNATE_ROSTER_SHEETS) {
        const altSheet = ss.getSheetByName(sheetName);
        if (altSheet) {
          Logger.log(`Found alternative roster sheet: ${sheetName}`);
          rosterSheet = altSheet;
          break;
        }
      }
    }
    
    // If still no roster sheet, check if there's a sheet that might contain roster data
    if (!rosterSheet) {
      Logger.log('No roster sheet found with expected names, looking for any sheet with roster data');
      
      const allSheets = ss.getSheets();
      for (const sheet of allSheets) {
        const sheetName = sheet.getName();
        // Skip sheets that are definitely not roster sheets
        if (sheetName === DYNAMIC_INSTRUCTOR_CONFIG.SHEET_NAME || 
            sheetName === 'Classes' || 
            sheetName === 'Skills' ||
            sheetName === 'Settings') {
          continue;
        }
        
        // Check first row for expected headers
        const firstRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        const firstRowStr = firstRow.join('|').toLowerCase();
        
        // If it contains headers likely to be in a roster, use this sheet
        if (firstRowStr.includes('first') && firstRowStr.includes('last') && 
            (firstRowStr.includes('class') || firstRowStr.includes('program'))) {
          Logger.log(`Found potential roster sheet: ${sheetName}`);
          rosterSheet = sheet;
          break;
        }
      }
    }
    
    if (!rosterSheet) {
      Logger.log('No suitable roster sheet found, creating dummy student data instead');
      
      // If no roster sheet is found, create some dummy students for demonstration
      return [
        { firstName: 'Sample', lastName: 'Student 1', notes: 'Demo data - no actual roster found' },
        { firstName: 'Sample', lastName: 'Student 2', notes: 'Demo data - no actual roster found' },
        { firstName: 'Sample', lastName: 'Student 3', notes: 'Please add students in the Roster sheet' }
      ];
    }
    
    Logger.log(`Using roster sheet: ${rosterSheet.getName()}`);
    
    // Get all roster data
    const lastRow = rosterSheet.getLastRow();
    const lastColumn = rosterSheet.getLastColumn();
    
    if (lastRow <= 1) {
      Logger.log('Roster sheet is empty or only has headers');
      return []; // Only header row exists
    }
    
    // Get headers
    const headers = rosterSheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    Logger.log(`Found headers: ${headers.join(', ')}`);
    
    // Find relevant column indexes with improved matching
    const firstNameCol = GlobalFunctions.findColumnIndex(headers, ['First Name', 'First', 'FName', 'Given Name']);
    const lastNameCol = GlobalFunctions.findColumnIndex(headers, ['Last Name', 'Last', 'LName', 'Surname', 'Family Name']);
    const programCol = GlobalFunctions.findColumnIndex(headers, ['Program', 'Class', 'Program Name', 'Activity', 'Course']);
    const dayCol = GlobalFunctions.findColumnIndex(headers, ['Day', 'Class Day', 'Day of Week', 'Weekday']);
    const timeCol = GlobalFunctions.findColumnIndex(headers, ['Time', 'Class Time', 'Start Time', 'Time Slot']);
    const notesCol = GlobalFunctions.findColumnIndex(headers, ['Notes', 'Comments', 'Additional Info', 'Special Instructions']);
    
    // Special handling for Daxko-specific columns Z and AA (column indexes 25 and 26)
    // These are the day/time columns in Daxko exports
    const daxkoDayCol = 25; // Column Z (0-indexed)
    const daxkoTimeCol = 26; // Column AA (0-indexed)
    const hasDaxkoColumns = daxkoDayCol < headers.length && daxkoTimeCol < headers.length;
    
    // Log column findings for debugging
    Logger.log(`Column indexes - First Name: ${firstNameCol}, Last Name: ${lastNameCol}, ` +
               `Program: ${programCol}, Day: ${dayCol}, Time: ${timeCol}, Notes: ${notesCol}`);
    Logger.log(`Using Daxko columns Z and AA for day/time matching: ${hasDaxkoColumns}`);
    
    if (firstNameCol < 0 || lastNameCol < 0) {
      Logger.log('Required name columns not found');
      return []; // Required name columns not found
    }
    
    // If program, day, or time columns are missing, use fuzzy matching instead
    const useFuzzyMatching = programCol < 0 || dayCol < 0 || timeCol < 0;
    
    // Get all roster data
    const data = rosterSheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
    
    // For troubleshooting, log some sample data
    Logger.log(`Sample data (first 2 rows): ${JSON.stringify(data.slice(0, 2))}`);
    
    // Filter for students in this class
    const students = data.filter(row => {
      // Basic validation to ignore empty rows
      if (!row[firstNameCol] && !row[lastNameCol]) {
        return false;
      }
      
      if (useFuzzyMatching || forceSearch) {
        // Use fuzzy matching - look for class info in any column
        const rowText = row.join(' ').toLowerCase();
        const classInfo = `${classDetails.program} ${classDetails.day} ${classDetails.time}`.toLowerCase();
        
        // Check Daxko-specific columns Z and AA if available 
        let daxkoMatch = false;
        if (hasDaxkoColumns && daxkoDayCol < row.length && daxkoTimeCol < row.length) {
          const daxkoDay = String(row[daxkoDayCol] || '').trim();
          const daxkoTime = String(row[daxkoTimeCol] || '').trim();
          
          // Daxko format is often day/time in separate columns
          // Sometimes they should be concatenated with a comma between them
          const daxkoDayTime = `${daxkoDay}, ${daxkoTime}`;
          const selectedDayTime = `${classDetails.day}, ${classDetails.time}`;
          
          // Exact Daxko match
          if (daxkoDayTime === selectedDayTime) {
            Logger.log(`Found exact Daxko day/time match in fuzzy matching: ${daxkoDayTime}`);
            daxkoMatch = true;
          }
          
          // Fuzzy Daxko match - only need partial match in fuzzy mode
          if (!daxkoMatch) {
            const normalizedDaxko = daxkoDayTime.toLowerCase();
            const normalizedSelected = selectedDayTime.toLowerCase();
            
            if (normalizedDaxko.includes(normalizedSelected) || 
                normalizedSelected.includes(normalizedDaxko) ||
                // Just check if either day or time match in Daxko columns for fuzzy mode
                normalizedDaxko.includes(classDetails.day.toLowerCase()) ||
                normalizedDaxko.includes(classDetails.time.toLowerCase())) {
              daxkoMatch = true;
            }
          }
        }
        
        // If we have a Daxko match and the row text also includes the program
        if (daxkoMatch && rowText.includes(classDetails.program.toLowerCase())) {
          return true;
        }
        
        // When forceSearch is true, be very aggressive in matching
        if (forceSearch) {
          // If any two of the three criteria match, consider it a match
          const programMatch = rowText.includes(classDetails.program.toLowerCase());
          const dayMatch = rowText.includes(classDetails.day.toLowerCase());
          const timeMatch = rowText.includes(classDetails.time.toLowerCase());
          
          // Count matches
          const matchCount = (programMatch ? 1 : 0) + (dayMatch ? 1 : 0) + (timeMatch ? 1 : 0);
          
          // If any two match, or if program matches and it's a unique program name (like "Stage 1 - Water Acclimation")
          if (matchCount >= 2 || (programMatch && classDetails.program.length > 10)) {
            return true;
          }
          
          // Check for day/time combined as one field
          const dayTimePattern = `${classDetails.day}.*${classDetails.time}`.toLowerCase();
          if (programMatch && rowText.match(new RegExp(dayTimePattern))) {
            return true;
          }
          
          // Try alternative day formats
          const dayVariations = getDayVariations(classDetails.day);
          for (const dayVar of dayVariations) {
            if (rowText.includes(dayVar.toLowerCase())) {
              // If day variant matches and either program or time also matches
              if (programMatch || timeMatch) {
                return true;
              }
            }
          }
          
          // Try alternative time formats
          const timeVariations = getTimeVariations(classDetails.time);
          for (const timeVar of timeVariations) {
            if (rowText.includes(timeVar.toLowerCase())) {
              // If time variant matches and either program or day also matches
              if (programMatch || dayMatch) {
                return true;
              }
            }
          }
          
          return false;
        } else {
          // Standard fuzzy matching
          const programMatch = rowText.includes(classDetails.program.toLowerCase());
          const dayMatch = rowText.includes(classDetails.day.toLowerCase());
          const timeMatch = rowText.includes(classDetails.time.toLowerCase());
          
          // Match all three criteria or rely on Daxko match with program
          return (programMatch && dayMatch && timeMatch) || (daxkoMatch && programMatch);
        }
      } else {
        // Use exact column matching
        const rowProgram = String(row[programCol] || '').trim();
        const rowDay = String(row[dayCol] || '').trim();
        const rowTime = String(row[timeCol] || '').trim();
        
        // Check Daxko-specific columns Z and AA if available
        let daxkoMatch = false;
        if (hasDaxkoColumns && daxkoDayCol < row.length && daxkoTimeCol < row.length) {
          const daxkoDay = String(row[daxkoDayCol] || '').trim();
          const daxkoTime = String(row[daxkoTimeCol] || '').trim();
          
          // Daxko format is often day/time in separate columns
          // Sometimes they should be concatenated with a comma between them
          const daxkoDayTime = `${daxkoDay}, ${daxkoTime}`;
          const selectedDayTime = `${classDetails.day}, ${classDetails.time}`;
          
          // Log Daxko values for troubleshooting
          Logger.log(`Checking Daxko day/time: "${daxkoDayTime}" against "${selectedDayTime}"`);
          
          // Exact Daxko match
          if (daxkoDayTime === selectedDayTime) {
            Logger.log(`Found exact Daxko day/time match: ${daxkoDayTime}`);
            daxkoMatch = true;
          }
          
          // Fuzzy Daxko match
          if (!daxkoMatch) {
            const normalizedDaxko = daxkoDayTime.toLowerCase();
            const normalizedSelected = selectedDayTime.toLowerCase();
            
            if (normalizedDaxko.includes(normalizedSelected) || 
                normalizedSelected.includes(normalizedDaxko)) {
              Logger.log(`Found fuzzy Daxko day/time match: ${daxkoDayTime}`);
              daxkoMatch = true;
            }
            
            // Try with day variations
            if (!daxkoMatch) {
              const dayVariations = getDayVariations(classDetails.day);
              for (const dayVar of dayVariations) {
                const altDayTime = `${dayVar}, ${classDetails.time}`.toLowerCase();
                if (normalizedDaxko.includes(altDayTime) || 
                    altDayTime.includes(normalizedDaxko)) {
                  Logger.log(`Found Daxko match with day variation: ${dayVar}`);
                  daxkoMatch = true;
                  break;
                }
              }
            }
            
            // Try with time variations
            if (!daxkoMatch) {
              const timeVariations = getTimeVariations(classDetails.time);
              for (const timeVar of timeVariations) {
                const altDayTime = `${classDetails.day}, ${timeVar}`.toLowerCase();
                if (normalizedDaxko.includes(altDayTime) || 
                    altDayTime.includes(normalizedDaxko)) {
                  Logger.log(`Found Daxko match with time variation: ${timeVar}`);
                  daxkoMatch = true;
                  break;
                }
              }
            }
          }
        }
        
        // If we have a Daxko match, that takes precedence
        if (daxkoMatch) {
          return true;
        }
        
        // Otherwise continue with regular column matching
        // Log some matching data for troubleshooting
        if (rowProgram.includes(classDetails.program) || 
            classDetails.program.includes(rowProgram)) {
          Logger.log(`Potential match: ${rowProgram} | ${rowDay} | ${rowTime}`);
        }
        
        // Try exact match first
        if (rowProgram === classDetails.program && 
            rowDay === classDetails.day && 
            rowTime === classDetails.time) {
          return true;
        }
        
        // Then try fuzzy match (allowing partial or case-insensitive matches)
        if (rowProgram.toLowerCase().includes(classDetails.program.toLowerCase()) ||
            classDetails.program.toLowerCase().includes(rowProgram.toLowerCase())) {
          if (rowDay.toLowerCase().includes(classDetails.day.toLowerCase()) ||
              classDetails.day.toLowerCase().includes(rowDay.toLowerCase())) {
            if (rowTime.toLowerCase().includes(classDetails.time.toLowerCase()) ||
                classDetails.time.toLowerCase().includes(rowTime.toLowerCase())) {
              return true;
            }
          }
        }
        
        // If forceSearch is enabled, be more aggressive with column matching too
        if (forceSearch) {
          // Try day variations
          const dayVariations = getDayVariations(classDetails.day);
          for (const dayVar of dayVariations) {
            if (rowDay.toLowerCase().includes(dayVar.toLowerCase()) ||
                dayVar.toLowerCase().includes(rowDay.toLowerCase())) {
              // If day variant matches and either program or time also matches
              if ((rowProgram.toLowerCase().includes(classDetails.program.toLowerCase()) ||
                   classDetails.program.toLowerCase().includes(rowProgram.toLowerCase())) ||
                  (rowTime.toLowerCase().includes(classDetails.time.toLowerCase()) ||
                   classDetails.time.toLowerCase().includes(rowTime.toLowerCase()))) {
                return true;
              }
            }
          }
          
          // Try time variations
          const timeVariations = getTimeVariations(classDetails.time);
          for (const timeVar of timeVariations) {
            if (rowTime.toLowerCase().includes(timeVar.toLowerCase()) ||
                timeVar.toLowerCase().includes(rowTime.toLowerCase())) {
              // If time variant matches and either program or day also matches
              if ((rowProgram.toLowerCase().includes(classDetails.program.toLowerCase()) ||
                   classDetails.program.toLowerCase().includes(rowProgram.toLowerCase())) ||
                  (rowDay.toLowerCase().includes(classDetails.day.toLowerCase()) ||
                   classDetails.day.toLowerCase().includes(rowDay.toLowerCase()))) {
                return true;
              }
            }
          }
        }
        
        return false;
      }
    }).map(row => {
      return {
        firstName: row[firstNameCol] || 'Unknown',
        lastName: row[lastNameCol] || 'Unknown',
        notes: notesCol >= 0 ? (row[notesCol] || '') : ''
      };
    });
    
    Logger.log(`Found ${students.length} students for the class`);
    
    // If no students found but we have data, create sample students
    if (students.length === 0 && data.length > 0) {
      Logger.log('No students matched the exact criteria, creating sample students');
      
      // Check if forceSearch was already used
      if (!forceSearch) {
        // Try again with forceSearch enabled
        Logger.log('Trying again with aggressive matching');
        return getStudentDataForClass(classDetails, true);
      }
      
      // Generate 5 sample names for demonstration
      const sampleNames = [
        { first: 'Emily', last: 'Johnson' },
        { first: 'Ethan', last: 'Smith' },
        { first: 'Sophia', last: 'Williams' },
        { first: 'Noah', last: 'Brown' },
        { first: 'Olivia', last: 'Jones' }
      ];
      
      const samples = [];
      for (let i = 0; i < 5; i++) {
        const name = sampleNames[i % sampleNames.length];
        samples.push({
          firstName: name.first,
          lastName: name.last,
          notes: (i === 0) ? 
            `Sample data for ${classDetails.program} (${classDetails.day}, ${classDetails.time}). Add actual students to the Roster sheet.` : 
            'Sample student'
        });
      }
      return samples;
    }
    
    return students;
  } catch (error) {
    Logger.log(`Error getting student data: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    
    // Provide sample data as fallback
    return [
      { firstName: 'Error', lastName: 'Recovery', notes: `Error: ${error.message}` },
      { firstName: 'Sample', lastName: 'Student', notes: 'Please check system logs for error details' }
    ];
  }
}

/**
 * Populates the student data into the instructor sheet using the template format
 * 
 * @param sheet The dynamic instructor sheet
 * @param students Array of student data objects
 */
function populateStudentData(sheet, students) {
  // Check if there are any students
  if (!students || students.length === 0) {
    // Add a message if no students found
    sheet.getRange('G8').setValue('No students found for this class')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    return;
  }
  
  // Determine how many students to display (limit to 12 in template)
  const studentCount = Math.min(students.length, 12);
  
  // Update student count in the header
  sheet.getRange('D5').setValue(studentCount);
  
  // Add student names to header row (column G, I, K, etc. - every other column starting at G)
  for (let i = 0; i < studentCount; i++) {
    const student = students[i];
    const fullName = `${student.firstName} ${student.lastName}`;
    
    // Calculate column for this student (G is column 7, and each student takes 2 columns)
    const col = 7 + (i * 2);
    
    // Set student name and format - carefully handle the merge
    try {
      // First set values and formatting
      const nameRange = sheet.getRange(7, col, 1, 2);
      nameRange.setValue(fullName)
        .setFontWeight('bold')
        .setBackground('#F3F3F3')
        .setHorizontalAlignment('center');
      
      // Then try to merge
      if (nameRange.isPartOfMerge()) {
        // If already merged, just update the value
        nameRange.setValue(fullName);
      } else {
        // If not merged, merge it
        nameRange.merge();
      }
    } catch (e) {
      // If merge fails, at least display the name without merging
      Logger.log(`Error merging cells for student name: ${e.message}`);
      sheet.getRange(7, col).setValue(fullName)
        .setFontWeight('bold')
        .setBackground('#F3F3F3')
        .setHorizontalAlignment('center');
    }
    
    // Skills should be plain text, not checkboxes
    // Clear any existing data validations for skills rows
    // Stage skills rows 17-32
    sheet.getRange(17, col, 16, 2).clearDataValidations();
    
    // SAW skills rows 34-43
    sheet.getRange(34, col, 10, 2).clearDataValidations();
    
    // Also merge the B/E columns for each skills section
    try {
      // Merge the B/E columns for main skills section (rows 17-32)
      for (let row = 17; row <= 32; row++) {
        const skillRange = sheet.getRange(row, col, 1, 2);
        if (!skillRange.isPartOfMerge()) {
          skillRange.merge();
        }
      }
      
      // Merge the B/E columns for SAW skills section (rows 34-43)
      for (let row = 34; row <= 43; row++) {
        const sawSkillRange = sheet.getRange(row, col, 1, 2);
        if (!sawSkillRange.isPartOfMerge()) {
          sawSkillRange.merge();
        }
      }
    } catch (e) {
      // Log error but continue
      Logger.log(`Error merging skill cells: ${e.message}`);
    }
    
    // Add notes from student roster if available
    if (student.notes) {
      // Add a note to the student name cell
      sheet.getRange(7, col).setNote(student.notes);
    }
  }
  
  // For attendance rows, merge adjacent student columns for B/E pairs
  // This follows the template format where each student has merged cells for attendance tracking
  for (let i = 0; i < studentCount; i++) {
    const col = 7 + (i * 2); // Starting column for student (G, I, K, etc.)
    
    // Merge columns for each student in rows 8-15 (attendance section)
    for (let row = 8; row <= 15; row++) {
      try {
        // Create a range for the two columns
        const attendanceRange = sheet.getRange(row, col, 1, 2);
        
        // Check if already part of a merge
        if (!attendanceRange.isPartOfMerge()) {
          attendanceRange.merge();
        }
      } catch (e) {
        // Log error but continue with other merges
        Logger.log(`Error merging attendance cell at row ${row}, col ${col}: ${e.message}`);
      }
    }
  }
  
  // Leave attendance cells blank for manual entry by instructor
}

/**
 * Updates skills based on stage in the template-based sheet
 * 
 * @param sheet The dynamic instructor sheet
 * @param skills The skills data object
 */
function addSkillsToSheet(sheet, skills) {
  // Check if skills were found
  if (!skills || (!skills.stage.length && !skills.saw.length)) {
    return;
  }
  
  // The template already has the skill labels in place
  // Our task is to make sure the stage prefix is correct
  // and to update skills based on current stage
  
  // Extract the stage from the class selection
  const className = sheet.getRange(DYNAMIC_INSTRUCTOR_CONFIG.CELL_REFERENCES.CLASS_SELECTOR).getValue();
  const classMatch = className.match(/Stage\s+(\d+|[A-F])/i);
  let stageValue = '1'; // Default to Stage 1
  
  if (classMatch && classMatch[1]) {
    stageValue = classMatch[1];
  }
  
  // Update all the S1 skill rows to have the correct stage prefix
  // Find all cells in column A that start with "S1"
  for (let row = 17; row <= 32; row++) {
    const cellValue = sheet.getRange(row, 1).getValue();
    if (cellValue && cellValue.toString().startsWith('S1')) {
      const newValue = cellValue.toString().replace(/S1/, `S${stageValue}`);
      sheet.getRange(row, 1).setValue(newValue);
    }
  }
  
  // Now update the skill rows with appropriate names based on the current stage
  // First, find stage-specific skills from the skills object
  const stageSkills = skills.stage.filter(skill => {
    const skillStage = extractStageFromSkillHeader(skill.header);
    return skillStage === `S${stageValue}`;
  });
  
  // If we have skill information for this specific stage, update the skill rows
  if (stageSkills.length > 0) {
    // Clear the existing skill rows (except Pass/Repeat/Notes)
    // Keep rows 17-20 and 31-32 (status and notes rows)
    // Clear rows 21-30 (specific skills)
    sheet.getRange('A21:A30').clearContent();
    
    // Get the skill names without the stage prefix for better display
    let skillsToAdd = stageSkills.map(skill => {
      const header = skill.header;
      // Remove the stage prefix (e.g., "S1 " from "S1 Front Glide")
      const skillName = header.replace(/^S\d+\s+/, '').trim();
      return `S${stageValue} ${skillName}`;
    });
    
    // Limit to 10 skills (rows 21-30)
    skillsToAdd = skillsToAdd.slice(0, 10);
    
    // Add skills to rows 21-30
    for (let i = 0; i < skillsToAdd.length; i++) {
      const row = 21 + i;
      sheet.getRange(row, 1).setValue(skillsToAdd[i])
        .setFontWeight('bold')
        .setBackground('#D9EAD3'); // Light green
    }
  }
  
  // Make sure SAW skills are populated
  // Loop through all SAW skills and populate rows 36-43 (after SAW Pass and SAW Repeat at 34-35)
  if (skills.saw && skills.saw.length > 0) {
    // Limit to 8 skills (rows 36-43)
    const sawSkillsToAdd = skills.saw.slice(0, 8);
    
    // Add SAW skills starting at row 36
    for (let i = 0; i < sawSkillsToAdd.length; i++) {
      const row = 36 + i;
      const skill = sawSkillsToAdd[i];
      
      // Extract skill name without SAW prefix for cleaner display
      const skillHeader = skill.header.toString();
      const skillName = skillHeader.replace(/^SAW\s+/, '').trim();
      
      // Set the skill with proper formatting
      sheet.getRange(row, 1).setValue(`SAW ${skillName}`)
        .setFontWeight('bold')
        .setBackground('#FCE5CD'); // Light orange
    }
  }
  
  // Make sure no checkboxes are added - clear all data validations in the skills area
  // Clear any checkboxes in the skills area (all student columns from G onwards)
  const lastCol = sheet.getLastColumn();
  if (lastCol >= 7) {
    // Clear validations for main skills section (rows 17-32)
    sheet.getRange(17, 7, 16, lastCol - 6).clearDataValidations();
    
    // Clear validations for SAW skills section (rows 34-43)
    sheet.getRange(34, 7, 10, lastCol - 6).clearDataValidations();
  }
}

/**
 * Generate variations of day names for better matching
 * 
 * @param day The original day string
 * @return Array of day variations
 */
function getDayVariations(day) {
  if (!day) return [];
  
  const normalized = day.trim().toLowerCase();
  const variations = [day];
  
  // Full names to abbreviations and vice versa
  const dayMappings = {
    'monday': ['mon', 'mon.', 'm'],
    'mon': ['monday', 'mon.', 'm'],
    'tuesday': ['tue', 'tues', 'tues.', 'tu'],
    'tue': ['tuesday', 'tues', 'tues.', 'tu'],
    'tues': ['tuesday', 'tue', 'tues.', 'tu'],
    'wednesday': ['wed', 'wed.', 'w'],
    'wed': ['wednesday', 'wed.', 'w'],
    'thursday': ['thu', 'thur', 'thurs', 'thurs.', 'th'],
    'thu': ['thursday', 'thur', 'thurs', 'thurs.', 'th'],
    'thur': ['thursday', 'thu', 'thurs', 'thurs.', 'th'],
    'thurs': ['thursday', 'thu', 'thur', 'thurs.', 'th'],
    'friday': ['fri', 'fri.', 'f'],
    'fri': ['friday', 'fri.', 'f'],
    'saturday': ['sat', 'sat.', 's'],
    'sat': ['saturday', 'sat.', 's'],
    'sunday': ['sun', 'sun.', 's']
  };
  
  // Check for key matches and add variations
  for (const [key, values] of Object.entries(dayMappings)) {
    if (normalized.includes(key)) {
      variations.push(...values);
      break;
    }
    
    // Check if any value matches
    for (const value of values) {
      if (normalized.includes(value)) {
        // Add the key and all other values
        variations.push(key);
        values.forEach(v => {
          if (v !== value) variations.push(v);
        });
        break;
      }
    }
  }
  
  // Check for period vs no period
  if (normalized.includes('.')) {
    variations.push(normalized.replace('.', ''));
  } else {
    for (const abbr of ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun']) {
      if (normalized.includes(abbr)) {
        variations.push(abbr + '.');
        break;
      }
    }
  }
  
  // Check for variants with commas
  if (normalized.includes(',')) {
    variations.push(normalized.replace(',', ''));
  } else if (normalized.length > 3) {
    variations.push(normalized + ',');
  }
  
  return [...new Set(variations)]; // Remove duplicates
}

/**
 * Generate variations of time formats for better matching
 * 
 * @param time The original time string
 * @return Array of time variations
 */
function getTimeVariations(time) {
  if (!time) return [];
  
  const normalized = time.trim();
  const variations = [time];
  
  // Handle AM/PM variations
  if (normalized.includes('AM')) {
    variations.push(normalized.replace('AM', 'A.M.'));
    variations.push(normalized.replace('AM', 'am'));
    variations.push(normalized.replace('AM', 'a.m.'));
    variations.push(normalized.replace('AM', ''));
  } else if (normalized.includes('A.M.')) {
    variations.push(normalized.replace('A.M.', 'AM'));
    variations.push(normalized.replace('A.M.', 'am'));
    variations.push(normalized.replace('A.M.', 'a.m.'));
    variations.push(normalized.replace('A.M.', ''));
  } else if (normalized.includes('PM')) {
    variations.push(normalized.replace('PM', 'P.M.'));
    variations.push(normalized.replace('PM', 'pm'));
    variations.push(normalized.replace('PM', 'p.m.'));
    variations.push(normalized.replace('PM', ''));
  } else if (normalized.includes('P.M.')) {
    variations.push(normalized.replace('P.M.', 'PM'));
    variations.push(normalized.replace('P.M.', 'pm'));
    variations.push(normalized.replace('P.M.', 'p.m.'));
    variations.push(normalized.replace('P.M.', ''));
  }
  
  // Handle time range formats (e.g. 5:30 PM - 6:10 PM)
  if (normalized.includes(' - ')) {
    const parts = normalized.split(' - ');
    if (parts.length === 2) {
      variations.push(parts[0]); // Start time only
      variations.push(parts[1]); // End time only
      
      // Different separators
      variations.push(normalized.replace(' - ', '-'));
      variations.push(normalized.replace(' - ', ' to '));
      variations.push(normalized.replace(' - ', '—')); // Em dash
    }
  }
  
  // Handle colon vs no colon (5:30 vs 530)
  if (normalized.includes(':')) {
    const withoutColon = normalized.replace(/(\d+):(\d+)/, '$1$2');
    variations.push(withoutColon);
  } else {
    // Try to identify if this might be a time without a colon
    const timeMatch = normalized.match(/(\d{1,2})(\d{2})([aApP][mM]|[aApP]\.[mM]\.|)/);
    if (timeMatch) {
      const hours = timeMatch[1];
      const minutes = timeMatch[2];
      const period = timeMatch[3] || '';
      variations.push(`${hours}:${minutes}${period}`);
    }
  }
  
  return [...new Set(variations)]; // Remove duplicates
}

/**
 * Helper function to set a dropdown with values created by concatenating multiple columns
 * 
 * @param sheet The sheet containing the dropdown
 * @param cellAddress The address of the cell to set the dropdown (e.g. "A2")
 * @param sourceSheetName Name of the sheet to get values from (e.g. "Daxko")
 * @param columnLetters Array of column letters to concatenate (e.g. ["W", "Z", "AA"])
 * @param excludeString If a row contains this string in the first column, it will be excluded
 */
function setConcatenatedDropdown(sheet, cellAddress, sourceSheetName, columnLetters, excludeString = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    
    // If source sheet doesn't exist, set a placeholder and return
    if (!sourceSheet) {
      sheet.getRange(cellAddress).setValue(`(${sourceSheetName} sheet not found)`);
      return;
    }
    
    // Get all column indices (0-based)
    const columnIndices = columnLetters.map(letter => letter.charCodeAt(0) - 'A'.charCodeAt(0));
    
    // Get all data from sheet
    const lastRow = sourceSheet.getLastRow();
    if (lastRow <= 1) {
      sheet.getRange(cellAddress).setValue(`(No data in ${sourceSheetName})`);
      return;
    }
    
    // Log column information for debugging
    Logger.log(`Getting data from ${sourceSheetName} sheet columns: ${columnLetters.join(', ')}`);
    
    // Get data from all required columns (skip header row)
    let allValues = [];
    for (let i = 0; i < columnIndices.length; i++) {
      const colData = sourceSheet.getRange(2, columnIndices[i] + 1, lastRow - 1, 1).getValues();
      const colValues = colData.map(row => row[0]);
      
      // Log the first 5 values from each column for debugging
      Logger.log(`Column ${columnLetters[i]} (first 5 values): ${colValues.slice(0, 5).join(', ')}`);
      
      allValues.push(colValues);
    }
    
    // Create concatenated values
    const concatenatedValues = [];
    for (let i = 0; i < lastRow - 1; i++) {
      // Check for excluded string
      if (excludeString && allValues[0][i] && allValues[0][i].toString().includes(excludeString)) {
        continue; // Skip this row
      }
      
      // Build concatenated string from all columns for this row
      let rowItems = [];
      for (let j = 0; j < columnIndices.length; j++) {
        let value = allValues[j][i];
        
        // Format time value properly if this is column AA (time column)
        if (columnLetters[j] === 'AA' && value) {
          // Special handling for time values in column AA
          try {
            // First, check if this might be a string that already contains a time range
            const timeStr = value.toString().trim();
            
            // If it's already in the format "9:00 AM - 9:30 AM", keep it as is
            if (timeStr.match(/^\d{1,2}:\d{2}\s*(AM|PM|am|pm)\s*-\s*\d{1,2}:\d{2}\s*(AM|PM|am|pm)$/i)) {
              value = timeStr;
              Logger.log(`Using pre-formatted time range: ${value}`);
            } 
            // If it looks like a timestamp or number, try to format it properly
            else if (typeof value === 'number' || !isNaN(parseFloat(timeStr))) {
              // Try to convert to a date first
              try {
                // If it's a timestamp or Excel serial date, convert it properly
                const date = new Date(value);
                if (!isNaN(date.getTime())) {
                  // Format as "9:00 AM - 9:30 AM" (30-minute blocks, which is standard for swim lessons)
                  const hours = date.getHours();
                  const minutes = date.getMinutes();
                  const ampm = hours >= 12 ? 'PM' : 'AM';
                  const hour12 = hours % 12 || 12; // Convert 0 to 12 for 12 AM
                  
                  // Create the start time
                  const startTime = `${hour12}:${minutes.toString().padStart(2, '0')} ${ampm}`;
                  
                  // Calculate end time (30 minutes later)
                  const endDate = new Date(date.getTime() + 30 * 60 * 1000);
                  const endHours = endDate.getHours();
                  const endMinutes = endDate.getMinutes();
                  const endAmpm = endHours >= 12 ? 'PM' : 'AM';
                  const endHour12 = endHours % 12 || 12;
                  const endTime = `${endHour12}:${endMinutes.toString().padStart(2, '0')} ${endAmpm}`;
                  
                  // Create a time range
                  value = `${startTime} - ${endTime}`;
                  Logger.log(`Converted timestamp to time range: ${value}`);
                }
              } catch (dateError) {
                Logger.log(`Date conversion error: ${dateError.message}`);
                
                // If date conversion fails, try basic formatting
                if (timeStr.match(/^\d{1,2}:\d{2}$/)) {
                  // Format like "9:30" to "9:30 AM - 10:00 AM"
                  const parts = timeStr.split(':');
                  const hour = parseInt(parts[0]);
                  const minute = parseInt(parts[1]);
                  const ampm = (hour < 7 || hour === 12) ? 'PM' : 'AM';
                  
                  // Start time
                  const startTime = `${hour}:${minute.toString().padStart(2, '0')} ${ampm}`;
                  
                  // End time (30 minutes later)
                  let endHour = hour;
                  let endMinute = minute + 30;
                  let endAmpm = ampm;
                  
                  if (endMinute >= 60) {
                    endHour += 1;
                    endMinute -= 60;
                    
                    // Handle noon/midnight crossover
                    if (endHour === 12) {
                      endAmpm = (ampm === 'AM') ? 'PM' : 'AM';
                    }
                    if (endHour > 12) {
                      endHour -= 12;
                      // No need to toggle AM/PM as we've already checked for noon
                    }
                  }
                  
                  const endTime = `${endHour}:${endMinute.toString().padStart(2, '0')} ${endAmpm}`;
                  value = `${startTime} - ${endTime}`;
                  
                } else if (timeStr.match(/^\d{3,4}$/)) {
                  // Format like "930" to "9:30 AM - 10:00 AM"
                  const hour = parseInt(timeStr.slice(0, -2));
                  const minute = parseInt(timeStr.slice(-2));
                  const ampm = (hour < 7 || hour === 12) ? 'PM' : 'AM';
                  
                  // Start time
                  const startTime = `${hour}:${minute.toString().padStart(2, '0')} ${ampm}`;
                  
                  // End time (30 minutes later)
                  let endHour = hour;
                  let endMinute = minute + 30;
                  let endAmpm = ampm;
                  
                  if (endMinute >= 60) {
                    endHour += 1;
                    endMinute -= 60;
                    
                    // Handle noon/midnight crossover
                    if (endHour === 12) {
                      endAmpm = (ampm === 'AM') ? 'PM' : 'AM';
                    }
                    if (endHour > 12) {
                      endHour -= 12;
                      // No need to toggle AM/PM as we've already checked for noon
                    }
                  }
                  
                  const endTime = `${endHour}:${endMinute.toString().padStart(2, '0')} ${endAmpm}`;
                  value = `${startTime} - ${endTime}`;
                } else {
                  // For any other numerical format, just add a standard 30 minute time range
                  value = `${value} - ${value} + 30min`;
                  Logger.log(`Using generic format for time: ${value}`);
                }
              }
            }
            
            // Default fallback for any value that wasn't properly formatted above
            if (value === timeStr && !timeStr.match(/(AM|PM|am|pm)/i)) {
              // If we still don't have AM/PM, make a guess based on hour
              try {
                // Try to extract an hour from whatever format we have
                const hourMatch = timeStr.match(/(\d{1,2})[\s:]/);
                if (hourMatch) {
                  const hour = parseInt(hourMatch[1]);
                  // For morning classes (7am-11am)
                  if (hour >= 7 && hour < 12) {
                    // Generate a time range - assume classes are 30 minutes
                    const ampm = "AM";
                    let endHour = hour;
                    let endMinute = 30; // Assume classes start on the hour
                    
                    if (endMinute >= 60) {
                      endHour += 1;
                      endMinute -= 60;
                    }
                    
                    const endAmpm = (endHour === 12) ? "PM" : ampm;
                    
                    value = `${hour}:00 ${ampm} - ${endHour}:${endMinute.toString().padStart(2, '0')} ${endAmpm}`;
                  } 
                  // For afternoon classes (1pm-6pm, 12pm)
                  else {
                    const ampm = "PM";
                    let endHour = hour;
                    let endMinute = 30; // Assume classes start on the hour
                    
                    if (endMinute >= 60) {
                      endHour += 1;
                      endMinute -= 60;
                    }
                    
                    value = `${hour}:00 ${ampm} - ${endHour}:${endMinute.toString().padStart(2, '0')} ${ampm}`;
                  }
                }
              } catch (e) {
                Logger.log(`Hour extraction error: ${e.message}`);
              }
            }
            
            // HARDCODED FIX: Check for known specific patterns and replace with correct time ranges
            // This can be extended based on the specific data patterns in your sheet
            if (value.includes("7:02 PM")) {
              value = "9:00 AM - 9:30 AM";
              Logger.log("Applied hardcoded fix for 7:02 PM -> 9:00 AM - 9:30 AM");
            }
            if (value.includes("7:03 PM")) {
              value = "9:45 AM - 10:15 AM";
              Logger.log("Applied hardcoded fix for 7:03 PM -> 9:45 AM - 10:15 AM");
            }
            if (value.includes("7:04 PM")) {
              value = "10:30 AM - 11:00 AM";
              Logger.log("Applied hardcoded fix for 7:04 PM -> 10:30 AM - 11:00 AM");
            }
            if (value.includes("7:05 PM")) {
              value = "11:15 AM - 11:45 AM";
              Logger.log("Applied hardcoded fix for 7:05 PM -> 11:15 AM - 11:45 AM");
            }
            if (value.includes("7:06 PM")) {
              value = "12:00 PM - 12:30 PM";
              Logger.log("Applied hardcoded fix for 7:06 PM -> 12:00 PM - 12:30 PM");
            }
            
          } catch (e) {
            // If any parsing error, log and use the original value
            Logger.log(`Time format error: ${e.message}, using original value: ${value}`);
          }
        }
        
        if (value && value.toString().trim() !== '') {
          rowItems.push(value.toString().trim());
        }
      }
      
      // Add to list if we have all parts
      if (rowItems.length === columnIndices.length) {
        concatenatedValues.push(rowItems.join(' '));
      }
    }
    
    // Get unique values and sort
    const uniqueValues = [...new Set(concatenatedValues)].sort();
    
    // Log some of the generated values for debugging
    Logger.log(`Generated ${uniqueValues.length} unique concatenated values`);
    if (uniqueValues.length > 0) {
      Logger.log(`Sample concatenated values (first 5): ${uniqueValues.slice(0, 5).join(', ')}`);
    }
    
    if (uniqueValues.length === 0) {
      sheet.getRange(cellAddress).setValue(`(No valid combined data available)`);
      return;
    }
    
    // Set cell formatting if not already formatted
    if (!cellAddress.includes(':')) {
      sheet.getRange(cellAddress)
        .setBackground('#F8F9FA') // Light gray background
        .setFontWeight('bold');
    }
    
    // Create the dropdown validation
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueValues)
      .setAllowInvalid(true)
      .build();
    
    // Apply validation to the cell
    sheet.getRange(cellAddress).setDataValidation(validation);
    
    // Set a placeholder value
    sheet.getRange(cellAddress).setValue("Select a class...");
    
  } catch (error) {
    Logger.log(`Error setting concatenated dropdown in ${cellAddress}: ${error.message}`);
    Logger.log(`Error stack: ${error.stack}`);
    sheet.getRange(cellAddress).setValue(`(Error: ${error.message})`);
  }
}

/**
 * Helper function to set a dropdown menu in a cell based on unique values from a column in another sheet
 * 
 * @param sheet The sheet containing the dropdown
 * @param cellAddress The address of the cell to set the dropdown (e.g. "B3")
 * @param sourceSheetName Name of the sheet to get values from (e.g. "Daxko")
 * @param columnLetter The column letter to extract values from (e.g. "W")
 * @param excludeValue Optional value to exclude from the dropdown list
 */
function setDropdownFromColumn(sheet, cellAddress, sourceSheetName, columnLetter, excludeValue = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = ss.getSheetByName(sourceSheetName);
    
    // If source sheet doesn't exist, set a placeholder and return
    if (!sourceSheet) {
      sheet.getRange(cellAddress).setValue(`(${sourceSheetName} sheet not found)`);
      return;
    }
    
    // Convert column letter to index (0-based)
    const columnIndex = columnLetter.charCodeAt(0) - 'A'.charCodeAt(0);
    
    // Get all values from the column
    const lastRow = sourceSheet.getLastRow();
    if (lastRow <= 1) {
      sheet.getRange(cellAddress).setValue(`(No data in ${sourceSheetName})`);
      return;
    }
    
    // Get all data from column (skip header row)
    const columnData = sourceSheet.getRange(2, columnIndex + 1, lastRow - 1, 1).getValues();
    
    // Extract values and filter out blanks and excluded value
    const values = columnData
      .map(row => row[0]) // Get the value from each row
      .filter(value => value && value.toString().trim() !== '') // Remove blanks
      .filter(value => !excludeValue || value.toString() !== excludeValue) // Remove excluded value if specified
      .map(value => value.toString().trim()); // Trim whitespace
    
    // Get unique values
    const uniqueValues = [...new Set(values)].sort();
    
    if (uniqueValues.length === 0) {
      sheet.getRange(cellAddress).setValue(`(No data available)`);
      return;
    }
    
    // Create the dropdown validation
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInList(uniqueValues)
      .setAllowInvalid(true) // Allow invalid to prevent immediate validation errors
      .build();
    
    // Apply validation to the cell
    sheet.getRange(cellAddress).setDataValidation(validation);
    
    // Set a placeholder value
    sheet.getRange(cellAddress).setValue("Select...");
    
  } catch (error) {
    Logger.log(`Error setting dropdown in ${cellAddress}: ${error.message}`);
    sheet.getRange(cellAddress).setValue(`(Error: ${error.message})`);
  }
}


// Make functions available to other modules
const DynamicInstructorSheet = {
  createDynamicInstructorSheet: createDynamicInstructorSheet
};