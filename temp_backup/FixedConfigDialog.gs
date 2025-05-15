/**
 * Fixed version of showConfigurationDialog that doesn't use insertTextBox
 * This function replaces the problematic code with a cell-based button alternative
 */
function showConfigurationDialogFixed() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let assumptionsSheet = ss.getSheetByName('Assumptions');
    
    if (!assumptionsSheet) {
      // If Assumptions sheet doesn't exist, create it
      prepareAssumptionsSheet(GlobalFunctions.safeGetProperty(CONFIG.SESSION_NAME) || '');
      assumptionsSheet = ss.getSheetByName('Assumptions');
    } else {
      // Update the Assumptions sheet with current values
      const sessionName = GlobalFunctions.safeGetProperty(CONFIG.SESSION_NAME) || '';
      const rosterFolderUrl = GlobalFunctions.safeGetProperty(CONFIG.ROSTER_FOLDER_URL) || '';
      const reportTemplateUrl = GlobalFunctions.safeGetProperty(CONFIG.REPORT_TEMPLATE_URL) || '';
      const swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL) || '';
      const parentHandbookUrl = GlobalFunctions.safeGetProperty(CONFIG.PARENT_HANDBOOK_URL) || '';
      const sessionProgramsUrl = GlobalFunctions.safeGetProperty(CONFIG.SESSION_PROGRAMS_URL) || '';
      
      assumptionsSheet.getRange('B7').setValue(sessionName);
      assumptionsSheet.getRange('B8').setValue(rosterFolderUrl);
      assumptionsSheet.getRange('B9').setValue(reportTemplateUrl);
      assumptionsSheet.getRange('B10').setValue(swimmerRecordsUrl);
      assumptionsSheet.getRange('B11').setValue(parentHandbookUrl);
      assumptionsSheet.getRange('B12').setValue(sessionProgramsUrl);
    }
    
    // Activate the Assumptions sheet to show configuration
    assumptionsSheet.activate();
    
    // Create a cell-based button instead of a textbox
    const buttonCell = assumptionsSheet.getRange(15, 2);
    assumptionsSheet.setRowHeight(15, 30);
    
    // Set cell formatting to make it look like a button
    buttonCell.setValue('Apply Configuration Changes')
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    
    // Add a note explaining how to use the button
    buttonCell.setNote('Click this cell and select "Apply Configuration Changes" from the YSL Hub menu to save these settings');
    
    // Add instructions below the button
    assumptionsSheet.getRange(16, 2).setValue('← Click this button after making changes')
      .setFontStyle('italic');
    
    ui.alert(
      'System Configuration',
      'You can update the configuration parameters in the Assumptions sheet. After making changes, ' +
      'click the blue "Apply Configuration Changes" button and then select the same option from the YSL Hub menu.\n\n' +
      'Note: Some changes may require restarting the system.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'showConfigurationDialogFixed', 
        'Error showing configuration dialog. Please try again or contact support.');
    } else {
      Logger.log(`Configuration dialog error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error showing configuration: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Creates a menu item for the fixed configuration dialog
 * This updates the createMenu function to use the fixed version
 */
function updateMenuForFixedDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('YSL Hub');
    
    menu.addItem('System Configuration (Fixed)', 'showConfigurationDialogFixed')
        .addItem('Apply Configuration Changes', 'AdministrativeModule_applyConfigurationChanges')
        .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog')
        .addToUi();
    
    ui.alert(
      'Fixed Configuration Menu Added',
      'A new menu item has been added to fix the configuration dialog issue. ' +
      'Please use "System Configuration (Fixed)" from the YSL Hub menu.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    Logger.log(`Error updating menu: ${error.message}`);
    return false;
  }
}