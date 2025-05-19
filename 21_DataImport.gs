/**
 * YSL Hub v2 Data Import Module
 * 
 * This module handles importing registration data and field mapping functionality.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-14
 */

/**
 * Imports registration data into the RegInfo sheet
 * Creates the sheet if it doesn't exist
 * Gets data from a user-provided Google Sheet URL or ID
 */
function importRegistrationData() {
  try {
    // Get the active spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Check if RegInfo sheet exists, create it if it doesn't
    let regInfoSheet = ss.getSheetByName('RegInfo');
    if (!regInfoSheet) {
      regInfoSheet = ss.insertSheet('RegInfo');
      setupNewSheet(regInfoSheet, 'RegInfo');
    }
    
    // Ask for source spreadsheet URL or ID
    const sourcePrompt = ui.prompt(
      'Import Registration Data',
      'Enter the URL or ID of the source Google Sheet:',
      ui.ButtonSet.OK_CANCEL
    );
    
    // Check if user canceled
    if (sourcePrompt.getSelectedButton() !== ui.Button.OK) {
      ui.alert('Import Canceled', 'Registration data import was canceled.', ui.ButtonSet.OK);
      return;
    }
    
    // Get the source ID or URL
    const sourceInput = sourcePrompt.getResponseText().trim();
    if (!sourceInput) {
      ui.alert('Invalid Input', 'Please provide a valid Google Sheet URL or ID.', ui.ButtonSet.OK);
      return;
    }
    
    // Extract the spreadsheet ID if a URL was provided
    let sourceId = sourceInput;
    if (sourceInput.includes('/')) {
      const match = sourceInput.match(/[-\w]{25,}/);
      if (match) {
        sourceId = match[0];
      }
    }
    
    // Try to open the source spreadsheet
    let sourceSpreadsheet;
    try {
      sourceSpreadsheet = SpreadsheetApp.openById(sourceId);
    } catch (error) {
      Logger.log('Failed to open source spreadsheet: ' + error);
      ui.alert(
        'Access Error',
        'Could not open the source spreadsheet. Make sure:\n\n' +
        '1. The URL or ID is correct\n' +
        '2. You have permission to access the spreadsheet\n' +
        '3. It is a Google Sheet (not Excel)',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Ask which sheet to import from
    const sourceSheets = sourceSpreadsheet.getSheets().map(sheet => sheet.getName());
    if (sourceSheets.length === 0) {
      ui.alert('Empty Spreadsheet', 'The source spreadsheet has no sheets.', ui.ButtonSet.OK);
      return;
    }
    
    // If there's only one sheet, use it, otherwise ask user to choose
    let selectedSheetName;
    if (sourceSheets.length === 1) {
      selectedSheetName = sourceSheets[0];
    } else {
      const sheetPrompt = ui.prompt(
        'Select Sheet',
        `The spreadsheet has multiple sheets. Enter the name of the sheet to import from:\n\n${sourceSheets.join(', ')}`,
        ui.ButtonSet.OK_CANCEL
      );
      
      // Check if user canceled
      if (sheetPrompt.getSelectedButton() !== ui.Button.OK) {
        ui.alert('Import Canceled', 'Registration data import was canceled.', ui.ButtonSet.OK);
        return;
      }
      
      selectedSheetName = sheetPrompt.getResponseText().trim();
      if (!sourceSheets.includes(selectedSheetName)) {
        ui.alert(
          'Invalid Sheet Name',
          `"${selectedSheetName}" was not found in the spreadsheet. Available sheets are:\n\n${sourceSheets.join(', ')}`,
          ui.ButtonSet.OK
        );
        return;
      }
    }
    
    // Get the selected sheet
    const sourceSheet = sourceSpreadsheet.getSheetByName(selectedSheetName);
    
    // Get data from source sheet
    const sourceData = sourceSheet.getDataRange().getValues();
    if (sourceData.length === 0) {
      ui.alert('Empty Sheet', 'The selected sheet has no data.', ui.ButtonSet.OK);
      return;
    }
    
    // Extract headers 
    const sourceHeaders = sourceData[0];
    
    // Confirm with the user
    const confirmResponse = ui.alert(
      'Confirm Import',
      `This will completely replace the RegInfo sheet with data from "${selectedSheetName}" including headers.\n\n` +
      `${sourceData.length - 1} rows and ${sourceHeaders.length} columns will be imported.\n\n` +
      'Do you want to continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse !== ui.Button.YES) {
      ui.alert('Import Canceled', 'Registration data import was canceled.', ui.ButtonSet.OK);
      return;
    }
    
    // Get the number of rows and columns in the source data
    const numRows = sourceData.length;
    const numColumns = sourceHeaders.length;
    
    Logger.log(`Source data dimensions: ${numRows} rows x ${numColumns} columns`);
    
    // Use a simple approach to avoid "out of bounds" errors
    try {
      // Get reference to the active spreadsheet
      const parentSpreadsheet = regInfoSheet.getParent();
      const sheetNames = parentSpreadsheet.getSheets().map(s => s.getName());
      
      // Create a temporary name that doesn't exist
      let tempName = "TempRegInfo";
      let counter = 1;
      while (sheetNames.includes(tempName)) {
        tempName = "TempRegInfo" + counter;
        counter++;
      }
      
      // Create a new temporary sheet with the required dimensions
      const tempSheet = parentSpreadsheet.insertSheet(tempName);
      
      // Make sure the new sheet has enough rows and columns
      if (tempSheet.getMaxRows() < numRows) {
        tempSheet.insertRowsAfter(1, numRows - tempSheet.getMaxRows());
      }
      
      if (tempSheet.getMaxColumns() < numColumns) {
        tempSheet.insertColumnsAfter(1, numColumns - tempSheet.getMaxColumns());
      }
      
      // Set column widths
      for (let i = 0; i < numColumns; i++) {
        tempSheet.setColumnWidth(i + 1, 120); // Default width
      }
      
      // Process in small chunks to avoid hitting limit errors
      const CHUNK_SIZE = 500; // Smaller chunks are safer
      
      for (let startRow = 0; startRow < numRows; startRow += CHUNK_SIZE) {
        const chunkSize = Math.min(CHUNK_SIZE, numRows - startRow);
        const chunk = sourceData.slice(startRow, startRow + chunkSize);
        
        // Write this chunk
        tempSheet.getRange(startRow + 1, 1, chunk.length, numColumns).setValues(chunk);
        
        // Small delay between chunks
        if (startRow + CHUNK_SIZE < numRows) {
          Utilities.sleep(50);
        }
      }
      
      // Format the header row
      tempSheet.getRange(1, 1, 1, numColumns)
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
      
      // Freeze the header row
      tempSheet.setFrozenRows(1);
      
      // Now delete the original sheet and rename the temp sheet
      parentSpreadsheet.deleteSheet(regInfoSheet);
      tempSheet.setName('RegInfo');
      regInfoSheet = tempSheet;
    } catch (sheetError) {
      Logger.log('Error recreating sheet: ' + sheetError);
      throw new Error('Error importing data: ' + sheetError.message);
    }
    
    // Set up field mapping dropdowns in YSLv6Hub sheet
    if (typeof FieldMapping !== 'undefined' && 
        typeof FieldMapping.setupFieldMappingControls === 'function') {
      FieldMapping.setupFieldMappingControls(sourceHeaders);
    } else {
      Logger.log('Warning: FieldMapping module not available. Field mapping controls not set up.');
    }
    
    // Activate the YSLv6Hub sheet to show the mapping controls
    const dashboardSheet = ss.getSheetByName('YSLv6Hub');
    if (dashboardSheet) {
      // Update the initialization status for Registration Data import
      if (typeof FieldMapping !== 'undefined' && 
          typeof FieldMapping.updateInitializationStatus === 'function') {
        FieldMapping.updateInitializationStatus(dashboardSheet, 'Import Registration Data', 'Complete');
      }
      
      dashboardSheet.activate();
    }
    
    // Show success message
    ui.alert(
      'Import Successful',
      `Successfully replaced RegInfo sheet with ${sourceData.length - 1} rows of data from "${selectedSheetName}".\n\nField mapping controls have been set up in the YSLv6Hub sheet.\n\nPlease complete the Session Name and Field Mappings in the YSLv6Hub sheet.`,
      ui.ButtonSet.OK
    );
    
    Logger.log(`Imported ${sourceData.length - 1} records from ${sourceId}/${selectedSheetName}`);
    
  } catch (error) {
    Logger.log('Error in importRegistrationData: ' + error);
    SpreadsheetApp.getUi().alert(
      'Import Failed',
      'Failed to import registration data: ' + error,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Sets up a new sheet with appropriate formatting and content
 * 
 * @param {Sheet} sheet - The sheet to set up
 * @param {string} sheetName - The name of the sheet
 */
function setupNewSheet(sheet, sheetName) {
  if (sheetName === 'RegInfo') {
    // Set up RegInfo sheet headers
    const headers = [
      'First Name', 'Last Name', 'Age', 'DOB', 'Guardian', 'Email', 
      'Phone', 'Class', 'Stage', 'Schedule', 'Notes', 'Start Date', 'End Date'
    ];
    
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#f3f3f3')
      .setFontWeight('bold');
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Set column widths
    for (let i = 0; i < headers.length; i++) {
      sheet.setColumnWidth(i + 1, 120);
    }
  }
}