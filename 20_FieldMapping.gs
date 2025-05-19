/**
 * YSL Hub v2 Field Mapping Module
 * 
 * This module handles field mapping between registration data and the system.
 * It manages the storage and retrieval of field mappings directly in the YSLv6Hub sheet.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-14
 */

/**
 * Sets up field mapping controls (dropdowns) in the YSLv6Hub sheet
 * @param {Array} headers - The headers from the RegInfo sheet
 */
function setupFieldMappingControls(headers) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('YSLv6Hub');
    
    if (!dashboardSheet) {
      Logger.log('YSLv6Hub sheet not found');
      return;
    }
    
    // Create or setup the dashboard with initialization status and field mappings
    setupYSLv6HubDashboard(dashboardSheet, headers);
    
    // Set up an onEdit trigger to save mappings when the user changes dropdown values
    const triggers = ScriptApp.getProjectTriggers();
    let hasEditTrigger = false;
    
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onEdit') {
        hasEditTrigger = true;
        break;
      }
    }
    
    if (!hasEditTrigger) {
      ScriptApp.newTrigger('onEdit')
        .forSpreadsheet(ss)
        .onEdit()
        .create();
    }
    
  } catch (error) {
    Logger.log('Error setting up field mapping controls: ' + error);
  }
}

/**
 * Sets up the YSLv6Hub dashboard with initialization status and field mappings
 * @param {Sheet} dashboardSheet - The YSLv6Hub sheet
 * @param {Array} headers - The headers from the RegInfo sheet
 */
function setupYSLv6HubDashboard(dashboardSheet, headers) {
  try {
    // Clear existing content
    dashboardSheet.clear();
    
    // Create dashboard header
    dashboardSheet.getRange('A1:C1').merge();
    dashboardSheet.getRange('A1').setValue('YSLv6Hub Dashboard')
      .setFontWeight('bold')
      .setFontSize(16)
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    // Section 1: Session Information
    dashboardSheet.getRange('A3:C3').merge();
    dashboardSheet.getRange('A3').setValue('Session Information')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#e9f2fe');
    
    // Session Name Input
    dashboardSheet.getRange('A4').setValue('Session Name (Required):')
      .setFontWeight('bold');
    dashboardSheet.getRange('B4').setValue('')
      .setBackground('#f0f0f0')
      .setNote('Enter the name of the current session (e.g., "Summer 2025")');
    
    // Section 2: System Initialization Status
    dashboardSheet.getRange('A6:C6').merge();
    dashboardSheet.getRange('A6').setValue('System Initialization Status')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#e9f2fe');
    
    // Status table headers
    dashboardSheet.getRange('A7:C7').setValues([['Action', 'Status', 'Notes']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Status table rows
    const statusRows = [
      ['Import Registration Data', 'Pending', 'Use "Import Registration Data" from menu'],
      ['Map Registration Fields', 'Pending', 'Complete field mapping below'],
      ['Generate SwimmerSkills', 'Pending', 'Use menu after field mapping'],
      ['Generate SwimmerLog', 'Pending', 'Use menu after field mapping'],
      ['Generate Groups Tracker', 'Pending', 'Use menu after field mapping'],
      ['Generate Communications Hub', 'Pending', 'Use menu after field mapping'],
      ['Generate Privates Tracker', 'Pending', 'Use menu after field mapping'],
      ['Sync Groups Tracker Info', 'Pending', 'Use menu after creating trackers'],
      ['Sync Privates Tracker Info', 'Pending', 'Use menu after creating trackers'],
      ['Session Transition', 'Pending', 'Use menu at end of session']
    ];
    
    dashboardSheet.getRange(8, 1, statusRows.length, 3).setValues(statusRows);
    
    // Format status cells
    for (let i = 0; i < statusRows.length; i++) {
      dashboardSheet.getRange(8 + i, 2).setBackground('#ffe0b2'); // Orange for pending
    }
    
    // Section 3: Field Mappings
    dashboardSheet.getRange('A19:C19').merge();
    dashboardSheet.getRange('A19').setValue('Field Mappings (Required)')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#e9f2fe');
    
    // Field mapping instructions
    dashboardSheet.getRange('A20:C20').merge();
    dashboardSheet.getRange('A20').setValue('Select the appropriate column from RegInfo for each field below:')
      .setFontStyle('italic');
    
    // Field mapping headers
    dashboardSheet.getRange('A21:C21').setValues([['Field Name', 'RegInfo Column', 'Description']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Field mapping rows - modified as requested
    const fieldMappingRows = [
      ['First Name (Required)', '', 'Student first name'],
      ['Last Name (Required)', '', 'Student last name'],
      ['Date of Birth', '', 'Student date of birth'],
      ['Program', '', 'Swim program name'],
      ['Class', '', 'Class or level information']
    ];
    
    dashboardSheet.getRange(22, 1, fieldMappingRows.length, 3).setValues(fieldMappingRows);
    
    // Format field name column
    dashboardSheet.getRange(22, 1, fieldMappingRows.length, 1).setFontWeight('bold');
    
    // Create dropdown validation for field mapping
    if (headers && headers.length > 0) {
      const headerValues = headers.filter(h => h != '');
      const headersList = SpreadsheetApp.newDataValidation()
        .requireValueInList(headerValues, true)
        .build();
      
      // Add validation to all mapping cells
      dashboardSheet.getRange(22, 2, fieldMappingRows.length, 1)
        .setDataValidation(headersList)
        .setBackground('#f0f0f0');
      
      // Get existing mappings
      const existingMappings = getFieldMappings();
      
      // Set existing values if available
      if (existingMappings) {
        const fieldKeys = ['firstNameCol', 'lastNameCol', 'dobCol', 'programCol', 'classCol'];
        const displayRows = [22, 23, 24, 25, 26]; // Rows in the sheet for each field
        
        fieldKeys.forEach((key, index) => {
          if (existingMappings[key]) {
            const colIndex = parseInt(existingMappings[key]);
            if (colIndex >= 0 && colIndex < headers.length && headers[colIndex]) {
              dashboardSheet.getRange(displayRows[index], 2).setValue(headers[colIndex]);
            }
          }
        });
      }
    }
    
    // Set column widths
    dashboardSheet.setColumnWidth(1, 200); // Field Name
    dashboardSheet.setColumnWidth(2, 200); // RegInfo Column
    dashboardSheet.setColumnWidth(3, 300); // Description
    
    // Add note about completion
    dashboardSheet.getRange('A28:C28').merge();
    dashboardSheet.getRange('A28').setValue('Complete all required fields before proceeding to create trackers')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#e6f2ff');
    
  } catch (error) {
    Logger.log('Error setting up YSLv6Hub dashboard: ' + error);
  }
}

/**
 * Handles edit events to save field mappings when the user changes dropdown values
 */
function onFieldMappingEdit(e) {
  try {
    // Check if the edit is in the YSLv6Hub sheet
    if (!e || !e.range || e.range.getSheet().getName() !== 'YSLv6Hub') {
      return;
    }
    
    const row = e.range.getRow();
    const col = e.range.getColumn();
    
    // Process session name edits (row 4, column B)
    if (row === 4 && col === 2) {
      const sessionName = e.value;
      if (sessionName) {
        // Update session name in properties
        const scriptProperties = PropertiesService.getScriptProperties();
        scriptProperties.setProperty('sessionName', sessionName);
        scriptProperties.setProperty('SESSION_NAME', sessionName);
        
        // Update initialization status
        updateInitializationStatus(e.range.getSheet(), 'Session Name', 'Complete');
        
        Logger.log(`Session name updated to: ${sessionName}`);
      }
      return;
    }
    
    // Only process field mapping edits in column B (dropdown selection column)
    if (col !== 2) {
      return;
    }
    
    // Check if this is one of the field mapping rows (rows 22-26)
    if (row < 22 || row > 26) {
      return;
    }
    
    // Get all the header values from RegInfo
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regInfoSheet = ss.getSheetByName('RegInfo');
    
    if (!regInfoSheet) {
      return;
    }
    
    const headers = regInfoSheet.getRange(1, 1, 1, regInfoSheet.getLastColumn()).getValues()[0];
    
    // Get the dropdown values
    const dashboardSheet = e.range.getSheet();
    
    // Map the row numbers to field names
    const displayRows = [22, 23, 24, 25, 26]; // Rows for First Name, Last Name, DOB, Program, Class
    const fieldNames = ['firstNameCol', 'lastNameCol', 'dobCol', 'programCol', 'classCol'];
    
    const mappings = {};
    
    // Read all current field mapping values from the sheet
    displayRows.forEach((displayRow, index) => {
      const value = dashboardSheet.getRange(displayRow, 2).getValue();
      if (value) {
        // Find the column index for this header value
        const colIndex = headers.findIndex(h => h === value);
        if (colIndex >= 0) {
          mappings[fieldNames[index]] = colIndex.toString();
        }
      }
    });
    
    // Save the mappings without showing an alert
    saveFieldMappingsQuiet(mappings);
    
    // Check if all required mappings are complete
    if (mappings.firstNameCol && mappings.lastNameCol) {
      updateInitializationStatus(dashboardSheet, 'Map Registration Fields', 'Complete');
    }
    
  } catch (error) {
    Logger.log('Error in onFieldMappingEdit: ' + error);
  }
}

/**
 * Updates the initialization status for a specific action
 * @param {Sheet} dashboardSheet - The YSLv6Hub sheet
 * @param {string} action - The action to update
 * @param {string} status - The new status ('Complete', 'Pending', or 'In Progress')
 */
function updateInitializationStatus(dashboardSheet, action, status) {
  try {
    // Find the row with the specified action
    const data = dashboardSheet.getDataRange().getValues();
    let actionRow = -1;
    
    for (let i = 7; i < data.length; i++) {
      if (data[i][0] === action) {
        actionRow = i + 1; // +1 because rows are 1-indexed
        break;
      }
    }
    
    // If action not found in the status table, check for "Map Registration Fields"
    if (actionRow === -1 && action === "Session Name") {
      for (let i = 7; i < data.length; i++) {
        if (data[i][0] === "Map Registration Fields") {
          actionRow = i + 1;
          break;
        }
      }
    }
    
    if (actionRow === -1) {
      return; // Action not found
    }
    
    // Update the status
    dashboardSheet.getRange(actionRow, 2).setValue(status);
    
    // Set background color based on status
    if (status === 'Complete') {
      dashboardSheet.getRange(actionRow, 2).setBackground('#c8e6c9'); // Green
    } else if (status === 'In Progress') {
      dashboardSheet.getRange(actionRow, 2).setBackground('#bbdefb'); // Blue
    } else {
      dashboardSheet.getRange(actionRow, 2).setBackground('#ffe0b2'); // Orange
    }
    
  } catch (error) {
    Logger.log(`Error updating initialization status: ${error}`);
  }
}

/**
 * Saves field mappings to the YSLv6Hub sheet without showing an alert
 */
function saveFieldMappingsQuiet(mappings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('YSLv6Hub');
    
    if (!dashboardSheet) {
      throw new Error('YSLv6Hub sheet not found');
    }
    
    // First, save to the visible field mapping section if it exists
    const dashboardData = dashboardSheet.getDataRange().getValues();
    let mappingSectionRow = -1;
    
    for (let i = 0; i < dashboardData.length; i++) {
      if (String(dashboardData[i][0]).includes('Field Mappings')) {
        mappingSectionRow = i;
        break;
      }
    }
    
    if (mappingSectionRow !== -1) {
      // Check if we have the new format with headers row
      let startRow = mappingSectionRow + 2; // Skip the headers row
      
      // Check if the new format has "Field Name" in the header row
      if (dashboardData[startRow - 1] && 
          String(dashboardData[startRow - 1][0]).includes('Field Name')) {
        
        // Get the RegInfo sheet to map column indices to headers
        const regInfoSheet = ss.getSheetByName('RegInfo');
        if (regInfoSheet) {
          const headers = regInfoSheet.getRange(1, 1, 1, regInfoSheet.getLastColumn()).getValues()[0];
          
          // Map internal field names to display field names
          const fieldDisplayMap = {
            'firstNameCol': 'First Name (Required)',
            'lastNameCol': 'Last Name (Required)',
            'dobCol': 'Date of Birth',
            'programCol': 'Program',
            'classCol': 'Class'
          };
          
          // Find rows for each field name
          Object.keys(mappings).forEach(field => {
            const displayName = fieldDisplayMap[field];
            if (displayName) {
              // Find the row with this display name
              for (let row = startRow; row < dashboardData.length; row++) {
                if (dashboardData[row][0] === displayName) {
                  // Get the header text for this column index
                  const colIndex = parseInt(mappings[field]);
                  if (!isNaN(colIndex) && colIndex >= 0 && colIndex < headers.length) {
                    // Set the dropdown value to the header text
                    dashboardSheet.getRange(row + 1, 2).setValue(headers[colIndex]);
                  }
                  break;
                }
              }
            }
          });
        }
      }
    }
    
    // Also save to hidden section for backward compatibility
    // Find or create the field mappings section (hidden)
    let hiddenMappingSectionRow = -1;
    
    for (let i = 0; i < dashboardData.length; i++) {
      if (dashboardData[i][0] === 'FIELD_MAPPINGS_DATA') {
        hiddenMappingSectionRow = i;
        break;
      }
    }
    
    // If hidden mapping section not found, create it at the bottom
    if (hiddenMappingSectionRow === -1) {
      hiddenMappingSectionRow = dashboardSheet.getLastRow() + 5;
      dashboardSheet.getRange(hiddenMappingSectionRow, 1).setValue('FIELD_MAPPINGS_DATA');
      
      // Hide this row
      dashboardSheet.hideRows(hiddenMappingSectionRow);
    }
    
    // Clear any existing hidden mappings
    const existingRows = dashboardSheet.getRange(hiddenMappingSectionRow + 1, 1, 10, 2);
    existingRows.clearContent();
    
    // Write the new mappings to hidden rows
    const mappingRows = [];
    
    Object.keys(mappings).forEach((field, index) => {
      const value = mappings[field];
      mappingRows.push([field, value]);
    });
    
    if (mappingRows.length > 0) {
      dashboardSheet.getRange(hiddenMappingSectionRow + 1, 1, mappingRows.length, 2)
        .setValues(mappingRows);
    }
    
    // Hide these rows too
    dashboardSheet.hideRows(hiddenMappingSectionRow + 1, mappingRows.length);
    
    return true;
  } catch (error) {
    Logger.log('Error saving field mappings quietly: ' + error);
    return false;
  }
}

/**
 * Gets field mappings from YSLv6Hub sheet or returns null if not set
 */
function getFieldMappings() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('YSLv6Hub');
    
    if (!dashboardSheet) {
      Logger.log('YSLv6Hub sheet not found');
      return null;
    }
    
    // Find the field mappings section in the dashboard
    // Search for the "Field Mappings" header
    const dashboardData = dashboardSheet.getDataRange().getValues();
    let mappingSectionRow = -1;
    
    for (let i = 0; i < dashboardData.length; i++) {
      if (String(dashboardData[i][0]).includes('Field Mappings')) {
        mappingSectionRow = i;
        break;
      }
    }
    
    // If mapping section not found, check for hidden data
    if (mappingSectionRow === -1) {
      // Try to find hidden mapping data
      for (let i = 0; i < dashboardData.length; i++) {
        if (dashboardData[i][0] === 'FIELD_MAPPINGS_DATA') {
          // Read from the next row
          let mappings = {};
          let row = i + 1;
          
          while (row < dashboardData.length && dashboardData[row][0] && dashboardData[row][1]) {
            const field = dashboardData[row][0];
            const value = dashboardData[row][1];
            mappings[field] = value;
            row++;
          }
          
          return mappings;
        }
      }
      
      Logger.log('Field Mappings section not found in YSLv6Hub sheet');
      return null;
    }
    
    // Check if we have the new format with headers row
    let startRow = mappingSectionRow + 2; // Skip the headers row
    
    // Check if the new format has "Field Name" in the header row
    if (dashboardData[startRow - 1] && 
        String(dashboardData[startRow - 1][0]).includes('Field Name')) {
      // New format: Field Name, RegInfo Column, Description
      const mappings = {};
      
      // Map from readable field names to internal field names
      const fieldMap = {
        'First Name (Required)': 'firstNameCol',
        'Last Name (Required)': 'lastNameCol',
        'Date of Birth': 'dobCol',
        'Program': 'programCol',
        'Class': 'classCol'
      };
      
      // Process each field mapping row (starting right after headers)
      for (let row = startRow; row < dashboardData.length; row++) {
        const fieldName = dashboardData[row][0];
        const value = dashboardData[row][1];
        
        // Stop if we encounter empty rows or end of section
        if (!fieldName || String(fieldName).trim() === '') {
          break;
        }
        
        // Map the readable field name to internal field name
        const internalName = fieldMap[fieldName];
        if (internalName && value) {
          // For dropdown values, we need to get the column index from RegInfo headers
          const regInfoSheet = ss.getSheetByName('RegInfo');
          if (regInfoSheet) {
            const headers = regInfoSheet.getRange(1, 1, 1, regInfoSheet.getLastColumn()).getValues()[0];
            const colIndex = headers.findIndex(h => h === value);
            if (colIndex >= 0) {
              mappings[internalName] = colIndex.toString();
            }
          }
        }
      }
      
      // Also get the session name
      const sessionNameRow = 4; // Row with session name input
      if (dashboardData.length > sessionNameRow && dashboardData[sessionNameRow-1][1]) {
        mappings['sessionName'] = dashboardData[sessionNameRow-1][1];
      }
      
      return mappings;
    } else {
      // Old format: direct field-value pairs
      const mappings = {};
      
      // Read until we hit an empty row or end of data
      while (startRow < dashboardData.length && 
             dashboardData[startRow][0] !== '' && 
             dashboardData[startRow][1] !== '') {
        const field = dashboardData[startRow][0];
        const value = dashboardData[startRow][1];
        
        if (field && value) {
          mappings[field] = value;
        }
        
        startRow++;
      }
      
      return mappings;
    }
  } catch (error) {
    Logger.log('Error getting field mappings: ' + error);
    return null;
  }
}

/**
 * Saves field mappings to the YSLv6Hub sheet
 */
function saveFieldMappings(mappings) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('YSLv6Hub');
    
    if (!dashboardSheet) {
      throw new Error('YSLv6Hub sheet not found');
    }
    
    // Find or create the field mappings section
    const dashboardData = dashboardSheet.getDataRange().getValues();
    let mappingSectionRow = -1;
    
    for (let i = 0; i < dashboardData.length; i++) {
      if (dashboardData[i][0] === 'Field Mappings') {
        mappingSectionRow = i;
        break;
      }
    }
    
    // If mapping section not found, create it at row 12
    if (mappingSectionRow === -1) {
      mappingSectionRow = 11; // 0-based index for row 12
      
      // Add the section header
      dashboardSheet.getRange('A12:C12').merge();
      dashboardSheet.getRange('A12').setValue('Field Mappings')
        .setFontWeight('bold')
        .setFontSize(14);
      
      // Add the column headers
      dashboardSheet.getRange('A13:C13').setValues([['Field Name', 'Column Index', 'Description']])
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
    }
    
    // Clear any existing mappings
    const existingRows = dashboardSheet.getRange(mappingSectionRow + 2, 1, 20, 3);
    existingRows.clearContent();
    
    // Write the new mappings
    const mappingRows = [];
    const fieldDescriptions = {
      'firstNameCol': 'First Name column',
      'lastNameCol': 'Last Name column',
      'dobCol': 'Date of Birth column',
      'ageCol': 'Age column',
      'classCol': 'Class column',
      'stageCol': 'Stage column'
    };
    
    Object.keys(mappings).forEach((field, index) => {
      const value = mappings[field];
      const description = fieldDescriptions[field] || '';
      
      mappingRows.push([field, value, description]);
    });
    
    if (mappingRows.length > 0) {
      dashboardSheet.getRange(mappingSectionRow + 2, 1, mappingRows.length, 3)
        .setValues(mappingRows);
    }
    
    SpreadsheetApp.getUi().alert(
      'Mappings Saved',
      'Field mappings have been saved successfully to the YSLv6Hub dashboard.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    Logger.log('Error saving field mappings: ' + error);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to save field mappings: ' + error,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}

/**
 * Shows the field mapping dialog
 */
function showFieldMappingDialog() {
  try {
    // Get the RegInfo sheet headers
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const regInfoSheet = ss.getSheetByName('RegInfo');
    
    if (!regInfoSheet) {
      SpreadsheetApp.getUi().alert(
        'RegInfo Missing',
        'The RegInfo sheet is required for field mapping. Please import registration data first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    
    // Get the first row (headers)
    const headers = regInfoSheet.getRange(1, 1, 1, regInfoSheet.getLastColumn()).getValues()[0];
    
    // Prepare header options for dropdown
    let headerOptions = '';
    headers.forEach((header, index) => {
      if (header) {
        headerOptions += `<option value="${index}">${header}</option>`;
      }
    });
    
    // Get existing mappings
    const existingMappings = getFieldMappings() || {};
    
    // Create the template and set variables
    const htmlTemplate = HtmlService.createTemplateFromFile('ui/html/FieldMapperDialog');
    htmlTemplate.headerOptions = headerOptions;
    htmlTemplate.existingMappings = JSON.stringify(existingMappings);
    
    const htmlOutput = htmlTemplate.evaluate()
      .setWidth(450)
      .setHeight(600)
      .setTitle('Map Registration Fields');
    
    // Show the dialog
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Map Registration Fields');
  } catch (error) {
    Logger.log('Error showing field mapping dialog: ' + error);
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to show field mapping dialog: ' + error,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Extracts students from RegInfo data for swimmer logs
 */
function extractStudentsForSwimmerLogs(regInfoData, fieldMappings) {
  try {
    const headers = regInfoData[0];
    const students = [];
    
    // Default to simple column detection if no mappings
    if (!fieldMappings) {
      // Simple fallback logic
      const headerIndices = {};
      headers.forEach((header, idx) => {
        const headerText = String(header).toLowerCase();
        if (headerText.includes('first')) headerIndices.firstName = idx;
        else if (headerText.includes('last')) headerIndices.lastName = idx;
        else if (headerText === 'dob' || headerText.includes('birth')) headerIndices.dob = idx;
        else if (headerText === 'age') headerIndices.age = idx;
      });
      
      // Process rows
      for (let i = 1; i < regInfoData.length; i++) {
        const row = regInfoData[i];
        const firstName = headerIndices.firstName !== undefined ? row[headerIndices.firstName] : '';
        const lastName = headerIndices.lastName !== undefined ? row[headerIndices.lastName] : '';
        const dob = headerIndices.dob !== undefined ? row[headerIndices.dob] : '';
        
        if (firstName || lastName) {
          students.push({
            firstName: firstName || '',
            lastName: lastName || '',
            dob: dob || '',
            fullName: `${firstName || ''} ${lastName || ''}`.trim()
          });
        }
      }
      
      return students;
    }
    
    // Use field mappings - parse column indices from strings to integers
    const firstNameColIdx = fieldMappings.firstNameCol ? parseInt(fieldMappings.firstNameCol) : -1;
    const lastNameColIdx = fieldMappings.lastNameCol ? parseInt(fieldMappings.lastNameCol) : -1;
    const dobColIdx = fieldMappings.dobCol ? parseInt(fieldMappings.dobCol) : -1;
    const ageColIdx = fieldMappings.ageCol ? parseInt(fieldMappings.ageCol) : -1;
    
    Logger.log(`Using field mappings - firstNameCol: ${firstNameColIdx}, lastNameCol: ${lastNameColIdx}, dobCol: ${dobColIdx}, ageCol: ${ageColIdx}`);
    
    // Process each row
    for (let i = 1; i < regInfoData.length; i++) {
      const row = regInfoData[i];
      
      const firstName = firstNameColIdx >= 0 ? row[firstNameColIdx] || '' : '';
      const lastName = lastNameColIdx >= 0 ? row[lastNameColIdx] || '' : '';
      const dob = dobColIdx >= 0 ? row[dobColIdx] || '' : '';
      const age = ageColIdx >= 0 ? row[ageColIdx] || '' : '';
      
      if (firstName || lastName) {
        students.push({
          firstName: firstName,
          lastName: lastName,
          dob: dob || age, // Use age as fallback for DOB
          fullName: `${firstName} ${lastName}`.trim()
        });
      }
    }
    
    return students;
  } catch (error) {
    Logger.log('Error extracting students for swimmer logs: ' + error);
    return [];
  }
}

// Make functions available to other modules
const FieldMapping = {
  setupFieldMappingControls: setupFieldMappingControls,
  setupYSLv6HubDashboard: setupYSLv6HubDashboard,
  onFieldMappingEdit: onFieldMappingEdit,
  updateInitializationStatus: updateInitializationStatus,
  saveFieldMappingsQuiet: saveFieldMappingsQuiet,
  getFieldMappings: getFieldMappings,
  saveFieldMappings: saveFieldMappings,
  showFieldMappingDialog: showFieldMappingDialog,
  extractStudentsForSwimmerLogs: extractStudentsForSwimmerLogs
};