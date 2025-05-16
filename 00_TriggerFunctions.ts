/**
 * YSL Hub v2 Trigger Functions
 * 
 * This file contains the global trigger functions that Google Apps Script
 * automatically executes in response to events.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-14
 */

/**
 * Global onOpen trigger function
 * Executed automatically when the spreadsheet is opened
 */
function onOpen() {
  try {
    // Try to create the full menu first
    createFullMenu();
    
    // Log successful menu creation
    Logger.log('Full menu created in onOpen trigger');
  } catch (error) {
    // If full menu fails, create a simple fallback menu
    Logger.log('Error creating full menu: ' + error.message);
    
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('YSL Hub')
      .addItem('Initialize System', 'AdministrativeModule_showInitializationDialog')
      .addItem('Fix Swimmer Records Access', 'fixSwimmerRecordsAccess_menuWrapper')
      .addItem('Fix Menu', 'fixMenu')
      .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog')
      .addToUi();
      
    Logger.log('Fallback menu created');
  }
}

/**
 * Emergency function to fix menu issues
 */
function fixMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Force properties to true
    PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    // Create the full menu
    createFullMenu();
    
    // Show confirmation
    ui.alert(
      'Menu Fixed',
      'System properties have been reset and menu has been created.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    Logger.log('Error fixing menu: ' + error.message);
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    return false;
  }
}

/**
 * Creates the full operational menu directly
 */
function createFullMenu() {
  const ui = SpreadsheetApp.getUi();
  
  const menu = ui.createMenu('YSL Hub');
  
  // 1. Class Management section - added directly to main menu instead of submenu
  menu.addItem('Generate Group Lesson Tracker', 'DynamicInstructorSheet_createDynamicInstructorSheet')
      .addSeparator()
      .addItem('Refresh Class List', 'DataIntegrationModule_updateClassSelector')
      .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData')
      .addSeparator();
    
  // 2. Communications section
  menu.addSubMenu(ui.createMenu('Communications')
    .addItem('Create Communications Hub', 'CommunicationModule_createCommunicationsHub')
    .addItem('Create Communication Log', 'CommunicationModule_createCommunicationLog')
    .addItem('Send Selected Communication', 'CommunicationModule_sendSelectedCommunication')
    .addSeparator()
    .addItem('Send Mid-Session Reports', 'ReportingModule_generateMidSessionReports')
    .addItem('Send End-Session Reports', 'ReportingModule_generateEndSessionReports')
    .addItem('Send Welcome Emails', 'CommunicationModule_sendWelcomeEmails'));
    
  // 3. System section
  menu.addSubMenu(ui.createMenu('System')
    .addItem('Create User Guide', 'UserGuide_createUserGuideSheet')
    .addSeparator()
    .addItem('View History', 'HistoryModule_createHistorySheet')
    .addSeparator()
    .addItem('Start New Session', 'SessionTransitionModule_startSessionTransition')
    .addItem('Resume Session Transition', 'SessionTransitionModule_resumeSessionTransition')
    .addSeparator()
    .addItem('System Configuration', 'AdministrativeModule_showConfigurationDialog')
    .addItem('Fix Swimmer Records Access', 'fixSwimmerRecordsAccess_menuWrapper')
    .addItem('Apply Configuration Changes', 'AdministrativeModule_applyConfigurationChanges')
    .addSeparator()
    .addItem('Show Logs', 'ErrorHandling_showLogViewer'));
    
  // Add About item and the menu to UI
  menu.addSeparator()
    .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog')
    .addToUi();
    
  return menu;
}

/**
 * Fixes Swimmer Records access issues by updating the configuration
 * with valid spreadsheet ID
 * This function is exposed globally for direct menu access
 */
function fixSwimmerRecordsAccess() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Show a dialog to get the correct URL from the user
    const result = ui.prompt(
      'Fix Swimmer Records Configuration',
      'Enter the URL or ID of the Swimmer Records spreadsheet:',
      ui.ButtonSet.OK_CANCEL
    );
    
    // Check if the user clicked "OK"
    if (result.getSelectedButton() === ui.Button.OK) {
      const userInput = result.getResponseText().trim();
      
      if (!userInput) {
        ui.alert('Error', 'You must enter a URL or ID.', ui.ButtonSet.OK);
        return;
      }
      
      // Try to validate the URL by attempting to open the spreadsheet
      try {
        // Extract ID from URL if needed
        let ssId = userInput;
        
        // If it looks like a URL, try to extract ID
        if (userInput.includes('/')) {
          const urlPattern = /[-\w]{25,}/;
          const match = userInput.match(urlPattern);
          if (match && match[0]) {
            ssId = match[0];
          }
        }
        
        // Try opening the spreadsheet to validate access
        const testSheet = SpreadsheetApp.openById(ssId);
        const sheetName = testSheet.getName(); // This will throw if access fails
        
        // Show confirmation that we could access the sheet
        ui.alert(
          'Access Confirmed',
          `Successfully accessed spreadsheet "${sheetName}" with ID: ${ssId}`,
          ui.ButtonSet.OK
        );
      } catch (accessError) {
        // Show warning but continue with saving
        ui.alert(
          'Warning',
          `Could not verify access to the spreadsheet. Error: ${accessError.message}\n\nThe URL will still be saved, but you may need to adjust sharing permissions.`,
          ui.ButtonSet.OK
        );
      }
      
      // Save the URL in script properties using both property keys to ensure it works
      PropertiesService.getScriptProperties().setProperty('swimmerRecordsUrl', userInput);
      PropertiesService.getScriptProperties().setProperty('SWIMMER_RECORDS_URL', userInput);
      
      // Also update it in the Assumptions sheet if it exists
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const assumptionsSheet = ss.getSheetByName('Assumptions');
        
        if (assumptionsSheet) {
          assumptionsSheet.getRange('B10').setValue(userInput);
        }
      } catch (e) {
        Logger.log(`Error updating Assumptions sheet: ${e.message}`);
        // Continue anyway
      }
      
      // Show success message
      ui.alert(
        'Configuration Updated',
        'Swimmer Records URL has been updated. Try rebuilding the instructor sheet now.',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    Logger.log(`Error fixing Swimmer Records access: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to update configuration: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Global onEdit trigger function
 * Executed automatically when the spreadsheet is edited
 * 
 * @param e - The edit event object
 */
function onEdit(e) {
  try {
    // If GlobalFunctions is available, use its onEdit function
    if (typeof GlobalFunctions !== 'undefined' && 
        typeof GlobalFunctions.onEdit === 'function') {
      GlobalFunctions.onEdit(e);
      return;
    }
    
    // Fallback implementation if GlobalFunctions is not available
    const sheet = e.source.getActiveSheet();
    const sheetName = sheet.getName();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();
    
    Logger.log(`Cell edited: ${sheetName} (${row}, ${col})`);

    // Instructor Sheet is now a static template
    // No dynamic functionality needed at this stage
  } catch (error) {
    Logger.log(`Error in onEdit trigger: ${error.message}`);
  }
}