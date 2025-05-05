/**
 * YSL Hub Instructor Resource Module - Modified
 * 
 * This module creates instructor sheets with different layouts for group and private lessons.
 * Group lessons have swimmers across the top and skills down the left.
 * Private lessons have date, student name, instructor, student record, and notes columns.
 * 
 * @author PenBay YMCA
 * @version 1.3
 */

/**
 * Configuration constants for module-specific settings
 */
const INSTRUCTOR_CONFIG = {
  // Format configuration
  SHEET_FORMAT: {
    HEADER_COLOR: '#4285F4',
    SECTION_COLOR: '#E0E0E0',
    SKILLS_COLOR: '#f3f3f3',
    TITLE_FONT_SIZE: 14,
    PAGE_WIDTH: 11,  // Landscape for wider tables
    PAGE_HEIGHT: 8.5
  },
  
  // Data validation options
  VALIDATION: {
    ASSESSMENT_OPTIONS: ['X', '/'],  // X = can perform, / = taught but cannot perform
    ATTENDANCE_OPTIONS: ['PRESENT', 'ABSENT', 'EXCUSED']
  },
  
  // Class types
  CLASS_TYPES: {
    GROUP: 'group',
    PRIVATE: 'private'
  }
};

/**
 * Generates instructor sheets for all selected classes.
 */
function generateInstructorSheets() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Retrieve selected classes information
    const selectedClasses = getSelectedClassesInfo();
    
    // Validate we have selected classes to process
    if (selectedClasses.length === 0) {
      ui.alert(
        'No Classes Selected',
        'Please select at least one class in the Classes sheet by setting the "Select Class" column to "Select".',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Process each selected class
    const results = processSelectedClasses(selectedClasses);
    
    // Show results summary
    showProcessingResults(results);
  } catch (error) {
    // Handle any unexpected errors
    Logger.log(`Error in generateInstructorSheets: ${error.message}`);
    ui.alert('Error', `An unexpected error occurred: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Retrieves information for all classes that have been selected in the Classes sheet.
 * 
 * @return {Array} Array of selected class objects with comprehensive class information
 */
function getSelectedClassesInfo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const classesSheet = ss.getSheetByName('Classes');
  
  if (!classesSheet) {
    throw new Error('Classes sheet not found. Please complete system initialization first.');
  }
  
  // Get all data from Classes sheet
  const classData = classesSheet.getDataRange().getValues();
  
  if (classData.length <= 1) {
    return [];
  }
  
  // Find selected classes (marked with 'Select' in column A)
  const selectedClasses = [];
  for (let i = 1; i < classData.length; i++) {
    if (classData[i][0] === 'Select') {
      // Check if this is a private lesson
      const isPrivate = classData[i][1].includes('Private');
      
      // Extract all class information as a structured object
      selectedClasses.push({
        rowIndex: i,
        program: classData[i][1],
        day: classData[i][2],
        // Store the full time string, not just the start time
        time: classData[i][3],
        location: classData[i][4],
        count: classData[i][5],
        instructor: classData[i][6],
        // Derive class type based on program name
        type: isPrivate ? INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE : INSTRUCTOR_CONFIG.CLASS_TYPES.GROUP
      });
    }
  }
  
  return selectedClasses;
}

/**
 * Extracts just the start time from a time string (e.g., "10:00 AM - 10:45 AM" -> "10:00 AM")
 * 
 * @param {string} timeString - The full time string with start and end times
 * @return {string} The start time only
 */
function extractStartTime(timeString) {
  if (!timeString) return '';
  
  // Check if there's a hyphen or dash indicating a time range
  const timeParts = timeString.split(/[-–—]/);
  if (timeParts.length > 1) {
    // Return just the start time and trim any whitespace
    return timeParts[0].trim();
  }
  
  // If no time range found, return the original string
  return timeString;
}

/**
 * Processes each selected class to generate appropriate instructor sheets.
 * 
 * @param {Array} selectedClasses - Array of selected class objects
 * @return {Object} Results object with success and error counts
 */
function processSelectedClasses(selectedClasses) {
  const results = {
    successCount: 0,
    errorCount: 0,
    errors: [] // Store specific errors for logging/diagnostics
  };
  
  selectedClasses.forEach(classInfo => {
    try {
      let sheet;
      
      // Choose the appropriate sheet format based on class type
      if (classInfo.type === INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE) {
        // For private lessons, use the private lesson format
        sheet = createPrivateLessonSheet(classInfo);
      } else {
        // For group classes, use the group class format
        sheet = createGroupClassSheet(classInfo);
      }
      
      if (sheet) {
        results.successCount++;
      }
    } catch (error) {
      results.errorCount++;
      results.errors.push({
        class: `${classInfo.program} (${classInfo.day}, ${classInfo.time})`,
        error: error.message
      });
      
      Logger.log(`Error generating sheet for ${classInfo.program} (${classInfo.day}, ${classInfo.time}): ${error.message}`);
    }
  });
  
  return results;
}

/**
 * Displays processing results to the user.
 * 
 * @param {Object} results - Results object with success and error counts
 */
function showProcessingResults(results) {
  const ui = SpreadsheetApp.getUi();
  
  ui.alert(
    'Sheet Generation Complete',
    `${results.successCount} instructor sheets generated successfully.\n` +
    (results.errorCount > 0 ? `${results.errorCount} sheets failed. Check logs for details.` : ''),
    ui.ButtonSet.OK
  );
}

/**
 * Creates a group class sheet with horizontal layout (swimmers across top, skills down left).
 * 
 * @param {Object} classInfo - Comprehensive class information object
 * @return {Sheet} The created instructor sheet
 */
function createGroupClassSheet(classInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Generate a standardized sheet name
  const sheetName = generateSheetName(classInfo);
  
  // Check if sheet already exists and handle replacement if needed
  let sheet = handleExistingSheet(sheetName);
  if (!sheet) {
    return null; // User opted not to replace existing sheet
  }
  
  try {
    // Retrieve roster and assessment criteria data
    const rosterData = retrieveRosterData(classInfo);
    const criteriaData = retrieveCriteriaForStage(classInfo.program);
    
    // Format the sheet with horizontal layout
    formatGroupClassSheet(sheet, classInfo, rosterData.roster, criteriaData);
    
    return sheet;
  } catch (error) {
    Logger.log(`Error in createGroupClassSheet: ${error.message}`);
    throw error;
  }
}

/**
 * Creates a private lesson sheet with date, student name, instructor, records and notes columns.
 * 
 * @param {Object} classInfo - Private lesson information
 * @return {Sheet} The created private lesson sheet
 */
function createPrivateLessonSheet(classInfo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Generate a standardized sheet name
  const sheetName = generateSheetName(classInfo);
  
  // Check if sheet already exists and handle replacement if needed
  let sheet = handleExistingSheet(sheetName);
  if (!sheet) {
    return null; // User opted not to replace existing sheet
  }
  
  // Format the sheet for private lessons
  formatPrivateLessonSheet(sheet, classInfo);
  
  return sheet;
}

/**
 * Generates a standardized sheet name based on class information.
 * 
 * @param {Object} classInfo - Class information object
 * @return {string} Standardized sheet name
 */
function generateSheetName(classInfo) {
  // Extract abbreviated stage from program name (e.g., "Stage 1" -> "S1")
  let stageAbbr;
  if (classInfo.program.includes('Stage')) {
    const match = classInfo.program.match(/Stage\s+([A-Za-z0-9]+)/i);
    if (match && match[1]) {
      stageAbbr = 'S' + match[1];
    } else {
      // Default if stage number can't be extracted
      stageAbbr = classInfo.program.substring(0, 2);
    }
  } else if (classInfo.program === 'Private Swim Lessons') {
    stageAbbr = 'PVT';
  } else {
    // For other program types, use first 2-3 characters
    stageAbbr = classInfo.program.substring(0, 3);
  }
  
  // Sanitize components to ensure valid sheet name
  const sanitizedDay = classInfo.day.replace(/[^a-zA-Z0-9]/g, '');
  
  // Extract just the start time for the sheet name
  const startTime = extractStartTime(classInfo.time);
  const sanitizedTime = startTime.replace(/[^a-zA-Z0-9]/g, '');
  
  return `${stageAbbr} ${sanitizedDay} ${sanitizedTime}`;
}

/**
 * Checks if a sheet already exists and handles replacement workflow.
 * 
 * @param {string} sheetName - Name of the sheet to check
 * @return {Sheet} New or existing sheet, or null if operation canceled
 */
function handleExistingSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (sheet) {
    // Ask if sheet should be replaced
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Sheet Already Exists',
      `An instructor sheet with this name already exists. Do you want to replace it?`,
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      ss.deleteSheet(sheet);
      sheet = null;
    } else {
      return null; // User chose not to replace
    }
  }
  
  // Create a new sheet if needed
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Retrieves roster data for a specific class.
 * 
 * @param {Object} classInfo - Class information object
 * @return {Object} Object containing roster data and headers
 */
function retrieveRosterData(classInfo) {
  try {
    // Log the search parameters for debugging
    Logger.log(`Searching for roster with: Program="${classInfo.program}", Day="${classInfo.day}", Time="${classInfo.time}"`);
    
    // Try to normalize the search parameters to increase match probability
    const normalizedProgram = classInfo.program.trim();
    const normalizedDay = classInfo.day.replace(/\.$/, '').trim(); // Remove trailing period
    
    // IMPORTANT: Use the full time string, not just the start time
    // This is critical for matching with the Daxko data
    const normalizedTime = classInfo.time.trim();
    
    Logger.log(`Normalized to: Program="${normalizedProgram}", Day="${normalizedDay}", Time="${normalizedTime}"`);
    
    // Pass the full normalized time string to getRosterForClass
    const rosterData = DataIntegrationModule.getRosterForClass(
      normalizedProgram,
      normalizedDay, 
      normalizedTime,
      classInfo.location
    );
    
    // For regular classes, we need students to be found
    if (classInfo.type !== INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE) {
      if (!rosterData || !rosterData.roster || rosterData.roster.length === 0) {
        throw new Error('No students found for this class');
      }
    }
    
    return rosterData;
  } catch (error) {
    // For private lessons, it's okay if no students are found - we'll create the sheet anyway
    if (classInfo.type === INSTRUCTOR_CONFIG.CLASS_TYPES.PRIVATE) {
      Logger.log('No students found for private lesson. Creating empty private lesson sheet.');
      return { roster: [], headers: [] };
    } else {
      // Log the error details for easier debugging
      Logger.log(`Error retrieving roster data: ${error.message}`);
      // For regular classes, rethrow the error
      throw error;
    }
  }
}

/**
 * Formats student names to show only first name and first initial of last name
 * 
 * @param {string} fullName - The student's full name
 * @return {string} Formatted name (e.g. "John D.")
 */
function formatStudentName(fullName) {
  if (!fullName) return '';
  
  const nameParts = fullName.trim().split(' ');
  if (nameParts.length === 1) return nameParts[0];
  
  const firstName = nameParts[0];
  const lastInitial = nameParts[nameParts.length - 1].charAt(0);
  
  return `${firstName} ${lastInitial}.`;
}

/**
 * Formats a group class sheet with horizontal layout (swimmers across top, skills down left).
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Object} classInfo - Class information object
 * @param {Array} roster - The class roster data
 * @param {Object} criteriaData - The assessment criteria for the class stage
 */
function formatGroupClassSheet(sheet, classInfo, roster, criteriaData) {
  // Set up header section
  createSheetHeader(sheet, 'INSTRUCTOR ASSESSMENT SHEET', roster.length);
  
  // Class information section
  createClassInfoSection(sheet, classInfo, roster.length);
  
  // Create horizontal roster with skills on left and students across top
  createHorizontalRosterSection(sheet, roster, criteriaData);
  
  // Format for printing
  formatSheetForPrinting(sheet);
}

/**
 * Creates the header section of an instructor sheet.
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {string} title - The title for the header
 * @param {number} studentCount - Number of students for header width
 */
function createSheetHeader(sheet, title, studentCount) {
  // Calculate the total width based on format (studentCount + 3 for skill name column)
  const totalColumns = Math.max(7, studentCount + 3); // Ensure at least 7 columns for small classes
  
  sheet.getRange(1, 1, 1, totalColumns).merge()
    .setValue(title)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.HEADER_COLOR)
    .setFontColor('white')
    .setFontSize(INSTRUCTOR_CONFIG.SHEET_FORMAT.TITLE_FONT_SIZE);
}

/**
 * Creates the class information section of an instructor sheet.
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Object} classInfo - Class information object
 * @param {number} studentCount - Number of students in the class
 */
function createClassInfoSection(sheet, classInfo, studentCount) {
  // Calculate the total width based on format
  const totalColumns = Math.max(7, studentCount + 3); // Ensure at least 7 columns for small classes
  
  // Section header
  sheet.getRange(2, 1, 1, totalColumns).merge()
    .setValue('Class Information')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR)
    .setHorizontalAlignment('center');
  
  // Row 1: Program/Stage and Instructor
  sheet.getRange(3, 1).setValue('Program:').setFontWeight('bold');
  sheet.getRange(3, 3, 1, 3).merge().setValue(classInfo.program);
  sheet.getRange(3, 7).setValue('Instructor:').setFontWeight('bold');
  sheet.getRange(3, 9, 1, 3).merge().setValue(classInfo.instructor);
  
  // Row 2: Day and Time
  sheet.getRange(4, 1).setValue('Day:').setFontWeight('bold');
  sheet.getRange(4, 3, 1, 3).merge().setValue(classInfo.day);
  sheet.getRange(4, 7).setValue('Time:').setFontWeight('bold');
  sheet.getRange(4, 9, 1, 3).merge().setValue(classInfo.time);
  
  // Row 3: Location and Student Count
  sheet.getRange(5, 1).setValue('Location:').setFontWeight('bold');
  sheet.getRange(5, 3, 1, 3).merge().setValue(classInfo.location);
  sheet.getRange(5, 7).setValue('Students:').setFontWeight('bold');
  sheet.getRange(5, 9, 1, 3).merge().setValue(studentCount);
}

/**
 * Formats a sheet for private lessons with completely revised layout.
 * Added extensive logging to diagnose data population issues.
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Object} classInfo - Private lesson information object
 */
function formatPrivateLessonSheet(sheet, classInfo) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Log the initial class info we're working with
    Logger.log(`===== PRIVATE LESSON SHEET - DEBUG =====`);
    Logger.log(`Class Info - Program: "${classInfo.program}", Day: "${classInfo.day}", Time: "${classInfo.time}"`);
    
    // Set up header section
    createSheetHeader(sheet, 'PRIVATE LESSON SHEET', 0);
    
    // Create basic info section
    sheet.getRange(2, 1, 1, 6).merge()
      .setValue('Lesson Information')
      .setFontWeight('bold')
      .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR)
      .setHorizontalAlignment('center');
    
    // Row 1: Instructor and Location
    sheet.getRange(3, 1).setValue('Instructor:').setFontWeight('bold');
    sheet.getRange(3, 2, 1, 3).merge().setValue(classInfo.instructor);
    
    // Row 2: Time and Location info
    sheet.getRange(4, 1).setValue('Day:').setFontWeight('bold');
    sheet.getRange(4, 2).setValue(classInfo.day);
    sheet.getRange(4, 3).setValue('Time:').setFontWeight('bold');
    sheet.getRange(4, 4).setValue(classInfo.time);
    sheet.getRange(4, 5).setValue('Location:').setFontWeight('bold');
    sheet.getRange(5, 1).setValue(classInfo.location);
    
    // Create lesson tracking section
    sheet.getRange(6, 1, 1, 6).merge()
      .setValue('LESSON TRACKING')
      .setFontWeight('bold')
      .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR)
      .setHorizontalAlignment('center');
    
    // Create column headers for lesson tracking with the new structure
    const headers = ['Date', 'Student Name', 'Instructor', 'Student Record', 'DivAb', 'Notes'];
    
    sheet.getRange(7, 1, 1, 6).setValues([headers])
      .setFontWeight('bold')
      .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SKILLS_COLOR);
    
    // Get swimmer records data for DivAb column
    let swimmerRecords = [];
    try {
      // Get the Swimmer Records URL from system configuration
      const config = AdministrativeModule.getSystemConfiguration();
      Logger.log(`Swimmer Records URL from config: ${config ? config.swimmerRecordsUrl : 'Not found'}`);
      
      if (config && config.swimmerRecordsUrl) {
        // Extract spreadsheet ID from URL
        const ssId = extractIdFromUrl(config.swimmerRecordsUrl);
        Logger.log(`Extracted Swimmer Records Spreadsheet ID: ${ssId}`);
        
        if (ssId) {
          // Open the Swimmer Records workbook
          const recordsWorkbook = SpreadsheetApp.openById(ssId);
          const swimmersSheet = recordsWorkbook.getSheetByName('Swimmers');
          
          if (swimmersSheet) {
            // Get all records including the DivAb column (DJ)
            const data = swimmersSheet.getDataRange().getValues();
            const headers = data[0];
            
            // Find column indices
            const nameColIndex = headers.indexOf('Name');
            const divAbColIndex = headers.indexOf('DivAb');
            
            Logger.log(`Swimmer Records - Name column index: ${nameColIndex}, DivAb column index: ${divAbColIndex}`);
            
            if (nameColIndex !== -1) {
              // Build lookup table for swimmer names to DivAb values
              for (let i = 1; i < data.length; i++) {
                const swimmerName = data[i][nameColIndex];
                const divAbValue = divAbColIndex !== -1 ? data[i][divAbColIndex] : '';
                
                if (swimmerName) {
                  swimmerRecords.push({
                    name: swimmerName,
                    divAb: divAbValue
                  });
                }
              }
              
              Logger.log(`Loaded ${swimmerRecords.length} swimmer records for DivAb lookup`);
              if (swimmerRecords.length > 0) {
                Logger.log(`Sample record - Name: "${swimmerRecords[0].name}", DivAb: "${swimmerRecords[0].divAb}"`);
              }
            }
          } else {
            Logger.log(`ERROR: Swimmers sheet not found in Swimmer Records workbook`);
          }
        }
      }
    } catch (error) {
      Logger.log(`ERROR loading Swimmer Records: ${error.message}`);
    }
    
    // Get session dates and student info from Daxko sheet
    try {
      // Attempt to get class roster and session dates
      const daxkoSheet = ss.getSheetByName('Daxko');
      
      if (!daxkoSheet) {
        Logger.log(`ERROR: Daxko sheet not found in the spreadsheet`);
        throw new Error('Daxko sheet not found');
      }
      
      // Get all data from Daxko sheet
      const daxkoData = daxkoSheet.getDataRange().getValues();
      const headers = daxkoData[0];
      
      // Find relevant column indices
      const programColIndex = headers.indexOf('Program');
      const dayColIndex = headers.indexOf('Day(s) of Week');
      const timeColIndex = headers.indexOf('Session Time');
      const firstNameColIndex = headers.indexOf('First Name');
      const lastNameColIndex = headers.indexOf('Last Name');
      const dobColIndex = headers.indexOf('DOB');
      
      // FIXED: Look for alternate column names that might contain session dates
      let segmentStartColIndex = headers.indexOf('Segment Start');
      
      // If Segment Start is not found, look for other date-related columns
      if (segmentStartColIndex === -1) {
        const possibleDateColumns = [
          'Start Date', 'Session Date', 'Class Date', 'Date', 
          'Session Start', 'Class Start', 'AB'  // AB is column index 27
        ];
        
        for (const columnName of possibleDateColumns) {
          const index = headers.indexOf(columnName);
          if (index !== -1) {
            segmentStartColIndex = index;
            Logger.log(`Using column "${columnName}" (index ${index}) for session dates`);
            break;
          }
        }
        
        // If still not found, try using column AB (index 27) directly
        if (segmentStartColIndex === -1) {
          segmentStartColIndex = 27; // Directly use column AB (0-based index 27)
          Logger.log(`Fallback: Using column AB (index 27) for session dates`);
        }
      }
      
      Logger.log(`Daxko Sheet Column Indices - Program: ${programColIndex}, Day: ${dayColIndex}, Time: ${timeColIndex}, First Name: ${firstNameColIndex}, Last Name: ${lastNameColIndex}, DOB: ${dobColIndex}, Segment Start: ${segmentStartColIndex}`);
      
      if (programColIndex === -1) {
        Logger.log(`ERROR: Program column not found in Daxko sheet. Headers found: ${headers.join(', ')}`);
        throw new Error('Required columns not found in Daxko sheet');
      }
      
      // Build the session information
      const sessions = [];
      const normalizedProgram = classInfo.program.trim();
      
      // FIXED: Don't remove period from day
      const normalizedDay = classInfo.day.trim();
      const normalizedTime = classInfo.time.trim();
      
      Logger.log(`Normalized search parameters - Program: "${normalizedProgram}", Day: "${normalizedDay}", Time: "${normalizedTime}"`);
      Logger.log(`Total rows in Daxko sheet: ${daxkoData.length - 1}`);
      
      // Filter for matching records
      let matchCount = 0;
      for (let i = 1; i < daxkoData.length; i++) {
        const row = daxkoData[i];
        
        if (!row[programColIndex]) {
          Logger.log(`Row ${i}: Skipping empty program`);
          continue;
        }
        
        const rowProgram = row[programColIndex].toString().trim();
        const rowDay = dayColIndex !== -1 && row[dayColIndex] ? row[dayColIndex].toString().trim() : '';
        const rowTime = timeColIndex !== -1 && row[timeColIndex] ? row[timeColIndex].toString().trim() : '';
        
        if (matchCount < 5) {
          Logger.log(`Row ${i} - Program: "${rowProgram}", Day: "${rowDay}", Time: "${rowTime}"`);
        }
        
        // FIXED: Modified matching logic to be more flexible
        const programMatch = rowProgram.trim() === normalizedProgram.trim() ||
                            rowProgram.includes(normalizedProgram) ||
                            normalizedProgram.includes(rowProgram);
        
        const dayMatch = (dayColIndex === -1) || 
                         (rowDay === normalizedDay) || 
                         (rowDay.replace('.', '') === normalizedDay.replace('.', ''));
        
        const timeMatch = (timeColIndex === -1) || 
                          (rowTime === normalizedTime) || 
                          (rowTime.trim() === normalizedTime.trim());
        
        // Log all matching components to debug
        if (programMatch && i < 50) {
          Logger.log(`Row ${i} - Program MATCH: "${rowProgram}" matches "${normalizedProgram}"`);
        }
        
        if (dayMatch && i < 50) {
          Logger.log(`Row ${i} - Day MATCH: "${rowDay}" matches "${normalizedDay}"`);
        }
        
        if (timeMatch && i < 50) {
          Logger.log(`Row ${i} - Time MATCH: "${rowTime}" matches "${normalizedTime}"`);
        }
        
        if (programMatch && dayMatch && timeMatch) {
          matchCount++;
          Logger.log(`ROW ${i} MATCHED! Program: "${rowProgram}", Day: "${rowDay}", Time: "${rowTime}"`);
          
          // Build student name
          let studentName = '';
          if (firstNameColIndex !== -1 && lastNameColIndex !== -1) {
            const firstName = row[firstNameColIndex] || '';
            const lastName = row[lastNameColIndex] || '';
            studentName = `${firstName} ${lastName}`.trim();
          }
          
          // Extract date
          let sessionDate = '';
          if (segmentStartColIndex !== -1) {
            const startDate = row[segmentStartColIndex];
            Logger.log(`Row ${i} - Date value from column ${segmentStartColIndex}: ${startDate}, type: ${typeof startDate}`);
            
            if (startDate instanceof Date) {
              sessionDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            } else if (startDate) {
              try {
                // Try parsing as a date if it's a string that looks like a date
                if (typeof startDate === 'string' && 
                    (startDate.includes('/') || startDate.includes('-'))) {
                  const dateObj = new Date(startDate);
                  if (!isNaN(dateObj.getTime())) {
                    sessionDate = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), 'MM/dd/yyyy');
                  } else {
                    sessionDate = startDate.toString();
                  }
                } else {
                  sessionDate = startDate.toString();
                }
              } catch (e) {
                sessionDate = startDate.toString();
              }
            }
          }
          
          // Extract DOB
          let dob = '';
          if (dobColIndex !== -1 && row[dobColIndex]) {
            const dobDate = row[dobColIndex];
            
            if (dobDate instanceof Date) {
              dob = Utilities.formatDate(dobDate, Session.getScriptTimeZone(), 'MM/dd/yyyy');
            } else {
              dob = dobDate.toString();
            }
          }
          
          Logger.log(`Match ${matchCount} - Student: "${studentName}", Date: "${sessionDate}", DOB: "${dob}"`);
          
          // FIXED: If there's no session date, create a dummy date based on the day of week
          if (!sessionDate) {
            // Try to create a date from the day of week
            const dayMap = {
              'Mon.': 'Monday',
              'Tues.': 'Tuesday',
              'Wed.': 'Wednesday',
              'Thurs.': 'Thursday',
              'Fri.': 'Friday',
              'Sat.': 'Saturday',
              'Sun.': 'Sunday'
            };
            
            const fullDay = dayMap[rowDay] || rowDay;
            sessionDate = `${fullDay} class`;
            Logger.log(`Created dummy date: "${sessionDate}" for day "${rowDay}"`);
          }
          
          // Add session
          sessions.push({
            date: sessionDate,
            studentName: studentName,
            dob: dob
          });
        }
      }
      
      Logger.log(`Found ${matchCount} matching records in Daxko sheet`);
      
      // FIXED: If no matches were found, create dummy entries for each day
      if (sessions.length === 0) {
        Logger.log(`No matches found. Creating dummy entries based on day of week.`);
        
        // Create 8 entries for this private lesson
        for (let i = 0; i < 8; i++) {
          const sessionDate = `Class ${i+1}`;
          sessions.push({
            date: sessionDate,
            studentName: '',
            dob: ''
          });
        }
      }
      
      // Sort sessions by date if possible
      sessions.sort((a, b) => {
        if (!a.date || !b.date) return 0;
        
        try {
          const dateA = new Date(a.date);
          const dateB = new Date(b.date);
          
          if (!isNaN(dateA.getTime()) && !isNaN(dateB.getTime())) {
            return dateA - dateB;
          } else {
            return a.date.localeCompare(b.date);
          }
        } catch (e) {
          return a.date.localeCompare(b.date);
        }
      });
      
      Logger.log(`Sorted sessions: ${sessions.length}`);
      
      // Add session data to the sheet
      const rowCount = Math.max(sessions.length, 10); // At least 10 rows
      Logger.log(`Creating ${rowCount} rows in private lesson sheet`);
      
      for (let i = 0; i < rowCount; i++) {
        // Set standard row height for all rows
        sheet.setRowHeight(8 + i, 21);
        
        // Set alternating row colors
        if (i % 2 === 0) {
          sheet.getRange(8 + i, 1, 1, 6).setBackground("#f3f3f3");
        }
        
        // Add session data if available
        if (i < sessions.length) {
          const session = sessions[i];
          
          if (session.date) {
            sheet.getRange(8 + i, 1).setValue(session.date);
            Logger.log(`Row ${i+8}, Column 1: Set date "${session.date}"`);
          }
          
          if (session.studentName) {
            const formattedName = formatStudentName(session.studentName);
            sheet.getRange(8 + i, 2).setValue(formattedName);
            Logger.log(`Row ${i+8}, Column 2: Set student name "${formattedName}"`);
            
            // Look up DivAb status
            let divAbFound = false;
            
            if (swimmerRecords.length > 0) {
              for (const record of swimmerRecords) {
                // Try to match the full name
                if (record.name === session.studentName) {
                  sheet.getRange(8 + i, 5).setValue(record.divAb);
                  divAbFound = true;
                  Logger.log(`Found DivAb value "${record.divAb}" for "${session.studentName}" (exact match)`);
                  break;
                }
              }
              
              // If not found by exact match, try more flexible matching
              if (!divAbFound) {
                for (const record of swimmerRecords) {
                  // Try to match by last name + first initial
                  const studentParts = session.studentName.split(' ');
                  const recordParts = record.name.split(' ');
                  
                  if (studentParts.length > 1 && recordParts.length > 1) {
                    const studentLastName = studentParts[studentParts.length - 1];
                    const studentFirstInitial = studentParts[0].charAt(0);
                    
                    const recordLastName = recordParts[recordParts.length - 1];
                    const recordFirstInitial = recordParts[0].charAt(0);
                    
                    if (studentLastName === recordLastName && studentFirstInitial === recordFirstInitial) {
                      sheet.getRange(8 + i, 5).setValue(record.divAb);
                      divAbFound = true;
                      Logger.log(`Found DivAb value "${record.divAb}" for "${session.studentName}" (partial match)`);
                      break;
                    }
                  }
                }
              }
              
              if (!divAbFound) {
                Logger.log(`No DivAb match found for "${session.studentName}"`);
              }
            }
          }
          
          if (session.dob) {
            sheet.getRange(8 + i, 4).setValue(`DOB: ${session.dob}`);
            Logger.log(`Row ${i+8}, Column 4: Set DOB "${session.dob}"`);
          }
        } else {
          Logger.log(`Row ${i+8}: No session data (empty row)`);
        }
      }
      
      Logger.log(`Added ${sessions.length} sessions to private lesson sheet`);
    } catch (error) {
      Logger.log(`ERROR populating private lesson dates: ${error.message}`);
      // Continue with empty sheet if we couldn't get session dates
      
      // Create default 10 empty rows
      for (let i = 0; i < 10; i++) {
        // Set standard row height
        sheet.setRowHeight(8 + i, 21);
        
        // Set alternating row colors
        if (i % 2 === 0) {
          sheet.getRange(8 + i, 1, 1, 6).setBackground("#f3f3f3");
        }
      }
    }
    
    // Set column widths
    sheet.setColumnWidth(1, 100);  // Date
    sheet.setColumnWidth(2, 150);  // Student Name
    sheet.setColumnWidth(3, 120);  // Instructor 
    sheet.setColumnWidth(4, 150);  // Student Record
    sheet.setColumnWidth(5, 80);   // DivAb
    sheet.setColumnWidth(6, 200);  // Notes
    
    // Set frozen rows to 7 only
    sheet.setFrozenRows(7);
    
    return sheet;
  } catch (error) {
    Logger.log(`CRITICAL ERROR in formatPrivateLessonSheet: ${error.message}`);
    throw error;
  }
}

/**
 * Retrieves assessment criteria for a specific program/stage.
 * Properly uses DataIntegrationModule to get criteria.
 * 
 * @param {string} program - Program or stage identifier
 * @return {Object} Object containing criteria data
 */
function retrieveCriteriaForStage(program) {
  // Log the program name for debugging
  Logger.log(`Getting criteria for stage: "${program}"`);
  
  // Convert program name to stage code manually
  let stageCode = '';
  
  if (program === 'Private Swim Lessons' || program.includes('Private')) {
    stageCode = 'private';
    Logger.log(`Converted to stage code: "${stageCode}"`);
  } else if (program.includes('Stage')) {
    // Extract stage letter/number from program name
    const match = program.match(/Stage\s+([A-Za-z0-9]+)/i);
    if (match && match[1]) {
      stageCode = 'S' + match[1];
      Logger.log(`Converted to stage code: "${stageCode}"`);
    }
  }
  
  try {
    // Get criteria data from DataIntegrationModule
    let criteriaData;
    if (stageCode) {
      criteriaData = DataIntegrationModule.getAssessmentCriteriaForStage(stageCode);
    } else {
      criteriaData = DataIntegrationModule.getAssessmentCriteriaForStage(program);
    }
    
    // Log what we received
    if (criteriaData && criteriaData.criteria) {
      Logger.log(`Found ${criteriaData.criteria.length} criteria for stage code ${stageCode}`);
    } else {
      Logger.log(`WARNING: No valid criteria returned for stage code ${stageCode}`);
    }
    
    return criteriaData;
  } catch (error) {
    Logger.log(`ERROR retrieving criteria: ${error.message}`);
    throw error;
  }
}

/**
 * Creates a horizontal roster section with skills on left and students across top.
 * Uses criteria from swimmer log and only hardcodes SAW skills.
 * 
 * @param {Sheet} sheet - The sheet to format
 * @param {Array} roster - The class roster data
 * @param {Object} criteriaData - Assessment criteria data
 */
function createHorizontalRosterSection(sheet, roster, criteriaData) {
  // Start row for roster section (after class info)
  const startRow = 7;
  
  // Get stage code to filter skills for that stage
  const stageCode = criteriaData.stageCode;
  Logger.log(`Creating horizontal roster with stage code: ${stageCode}`);
  
  // Create section header - make it the same size as other headers
  sheet.getRange(startRow, 1, 1, 20 + roster.length).merge()
    .setValue('STUDENT ASSESSMENT')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR)
    .setHorizontalAlignment('center');
  
  // Add assessment key
  sheet.getRange(startRow + 1, 1, 1, 20 + roster.length).merge()
    .setValue('KEY: X = Swimmer can perform the skill   / = Swimmer was taught the skill but cannot yet perform it')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Create student header row
  const studentRow = startRow + 2;
  
  // Add column for skill area in a merged cell
  sheet.getRange(studentRow, 1, 1, 3).merge()
    .setValue('SKILLS')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SKILLS_COLOR)
    .setHorizontalAlignment('center');
  
  // Add beginning/end columns - new columns at positions 4 and 5
  sheet.getRange(studentRow, 4).setValue('B')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SKILLS_COLOR)
    .setHorizontalAlignment('center');
  
  sheet.getRange(studentRow, 5).setValue('E')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SKILLS_COLOR)
    .setHorizontalAlignment('center');
  
  // Add student names across the top starting from column 20
  for (let i = 0; i < roster.length; i++) {
    const student = roster[i];
    // Format student name to first name and last initial
    const formattedName = formatStudentName(student.name);
    
    sheet.getRange(studentRow, i + 20).setValue(formattedName)
      .setFontWeight('bold')
      .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SKILLS_COLOR)
      .setTextRotation(90); // Rotate student names
  }
  
  // Add skills rows
  let currentRow = studentRow + 1;
  
  // Add attendance tracking section first
  sheet.getRange(currentRow, 1, 1, 19).merge()
    .setValue('ATTENDANCE')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 21);
  
  // Format student cells for the attendance header row
  for (let i = 0; i < roster.length; i++) {
    sheet.getRange(currentRow, i + 20)
      .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR);
  }
  
  currentRow++;
  
  // Add 8 class sessions for attendance - without dropdowns
  for (let i = 1; i <= 8; i++) {
    sheet.getRange(currentRow, 1, 1, 19).merge()
      .setValue(`Class ${i}`)
      .setHorizontalAlignment('center');
    
    // Set alternating row colors
    if (i % 2 === 0) {
      sheet.getRange(currentRow, 1, 1, 19 + roster.length).setBackground("#f3f3f3");
    }
    
    // Set standard row height
    sheet.setRowHeight(currentRow, 21);
    
    currentRow++;
  }
  
  // Add a separator row
  sheet.getRange(currentRow, 1, 1, 19 + roster.length)
    .setBackground('#CCCCCC');
  sheet.setRowHeight(currentRow, 21);
  currentRow++;
  
  // Add skills assessment section
  sheet.getRange(currentRow, 1, 1, 19).merge()
    .setValue('SKILLS ASSESSMENT')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR)
    .setHorizontalAlignment('center');
  sheet.setRowHeight(currentRow, 21);
  
  // Format student cells for the skills header row
  for (let i = 0; i < roster.length; i++) {
    sheet.getRange(currentRow, i + 20)
      .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR);
  }
  
  currentRow++;
  
  // Process stage-specific skills from criteria data
  let rowCounter = 0;
  
  // Extract stage-specific criteria headers from criteriaData
  const stageHeaders = [];
  
  // Check if we have valid criteria from swimmer log
  if (criteriaData && criteriaData.criteria && criteriaData.criteria.length > 0) {
    criteriaData.criteria.forEach(criterion => {
      if (criterion.header && criterion.header.startsWith(stageCode)) {
        stageHeaders.push(criterion.header);
        Logger.log(`Added header from criteria: ${criterion.header}`);
      }
    });
  } else if (criteriaData && criteriaData.allHeaders && criteriaData.allHeaders.length > 0) {
    // Try using allHeaders as fallback
    criteriaData.allHeaders.forEach(header => {
      if (header && header.startsWith(stageCode)) {
        stageHeaders.push(header);
        Logger.log(`Added header from allHeaders: ${header}`);
      }
    });
  }
  
  Logger.log(`Found ${stageHeaders.length} stage-specific headers`);
  
  // Process headers if we have them
  if (stageHeaders.length > 0) {
    // Process all stage-specific skills
    for (const header of stageHeaders) {
      // Determine if this header should have validation
      const hasValidation = !header.includes('Notes') && !header.includes('Sent');
      
      // Merge columns for skill name
      sheet.getRange(currentRow, 1, 1, 3).merge()
        .setValue(header)
        .setFontWeight('bold')
        .setHorizontalAlignment('center');
      
      // Apply alternating row colors
      if (rowCounter % 2 === 0) {
        sheet.getRange(currentRow, 1, 1, 19 + roster.length).setBackground("#f3f3f3");
      }
      
      // Set standard row height
      sheet.setRowHeight(currentRow, 21);
      
      // Add validation for B and E columns if needed
      if (hasValidation) {
        // Set up validation for B column
        const skillValidation = SpreadsheetApp.newDataValidation()
          .requireValueInList(INSTRUCTOR_CONFIG.VALIDATION.ASSESSMENT_OPTIONS, true)
          .build();
          
        sheet.getRange(currentRow, 4).setDataValidation(skillValidation);
        sheet.getRange(currentRow, 5).setDataValidation(skillValidation);
        
        // Set up validation for each student
        for (let j = 0; j < roster.length; j++) {
          sheet.getRange(currentRow, j + 20).setDataValidation(skillValidation);
        }
      }
      
      currentRow++;
      rowCounter++;
    }
  } else {
    // Add a message if no criteria were found
    sheet.getRange(currentRow, 1, 1, 19).merge()
      .setValue(`No criteria found for ${stageCode}. Run "Pull Latest Assessment Criteria" from the menu.`)
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    currentRow++;
  }
  
  // Add a separator row
  sheet.getRange(currentRow, 1, 1, 19 + roster.length)
    .setBackground('#CCCCCC');
  sheet.setRowHeight(currentRow, 21);
  currentRow++;
  
  // Add Safety Around Water section header
  sheet.getRange(currentRow, 1, 1, 19).merge()
    .setValue('SAFETY AROUND WATER CRITERIA')
    .setFontWeight('bold')
    .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR)
    .setHorizontalAlignment('center');
  
  // Format student cells for the SAW header row
  for (let i = 0; i < roster.length; i++) {
    sheet.getRange(currentRow, i + 20)
      .setBackground(INSTRUCTOR_CONFIG.SHEET_FORMAT.SECTION_COLOR);
  }
  
  sheet.setRowHeight(currentRow, 21);
  currentRow++;
  
  // Add standard Safety Around Water skills (this is the only hardcoded part)
  const sawSkills = [
    'SAW Submerge Face',
    'SAW Submerge, bob independently',
    'SAW Front glide, 5 ft., exit',
    'SAW Back float, 10 sec., roll, front glide, exit',
    'Swim, float, swim, 10 ft.',
    'Jump, independently',
    'Jump, push, turn, grab, assisted',
    'Jump, push, turn, grab'
  ];
  
  // Process all SAW skills
  for (let i = 0; i < sawSkills.length; i++) {
    // Merge columns for SAW skill name
    sheet.getRange(currentRow, 1, 1, 3).merge()
      .setValue(sawSkills[i])
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    // Apply alternating row colors
    if (i % 2 === 0) {
      sheet.getRange(currentRow, 1, 1, 19 + roster.length).setBackground("#f3f3f3");
    }
    
    // Set standard row height
    sheet.setRowHeight(currentRow, 21);
    
    // Add validation for B and E columns
    const skillValidation = SpreadsheetApp.newDataValidation()
      .requireValueInList(INSTRUCTOR_CONFIG.VALIDATION.ASSESSMENT_OPTIONS, true)
      .build();
      
    sheet.getRange(currentRow, 4).setDataValidation(skillValidation);
    sheet.getRange(currentRow, 5).setDataValidation(skillValidation);
    
    // Set up validation for each student
    for (let j = 0; j < roster.length; j++) {
      sheet.getRange(currentRow, j + 20).setDataValidation(skillValidation);
    }
    
    currentRow++;
  }
  
  // Add borders to separate student columns
  for (let i = 0; i < roster.length; i++) {
    const col = i + 20;
    const range = sheet.getRange(studentRow, col, currentRow - studentRow, 1);
    range.setBorder(true, true, true, true, false, false);
  }
  
  // Set column widths
  sheet.setColumnWidth(1, 50);  // First column of skill name
  sheet.setColumnWidth(2, 50);  // Second column of skill name
  sheet.setColumnWidth(3, 50);  // Third column of skill name
  sheet.setColumnWidth(4, 30);  // B column
  sheet.setColumnWidth(5, 30);  // E column
  
  // Set remaining columns to standard width
  for (let i = 6; i < 20; i++) {
    sheet.setColumnWidth(i, 15);  // Spacer columns
  }
  
  // Set student columns width
  for (let i = 0; i < roster.length; i++) {
    sheet.setColumnWidth(i + 20, 50);  // Student columns
  }
  
  Logger.log(`Horizontal roster section created with ${roster.length} students`);
}

/**
 * Formats a sheet for optimal printing on landscape 11x8.5 paper.
 * 
 * @param {Sheet} sheet - The sheet to format
 */
function formatSheetForPrinting(sheet) {
  try {
    // For group lesson sheets, freeze header rows to row 9 to show students
    sheet.setFrozenRows(9);
    
    // Set print settings
    sheet.setRowHeight(1, 30); // Make header row taller
  } catch (error) {
    Logger.log("Error in formatSheetForPrinting: " + error.message);
  }
}

/**
 * Utility function to extract ID from a Google Drive URL.
 * 
 * @param {string} url - The Google Drive URL
 * @return {string|null} The extracted ID, or null if not found
 */
function extractIdFromUrl(url) {
  if (!url) return null;
  
  // Extract ID from various URL formats
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
InstructorResourceModule = {
  generateInstructorSheets: generateInstructorSheets
};