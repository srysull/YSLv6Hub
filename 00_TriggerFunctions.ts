/**
 * YSL Hub v2 Trigger Functions
 * 
 * This file contains additional trigger functions for Google Apps Script.
 * The main onOpen trigger is now in 00_MenuSystem.ts.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-18
 */

/**
 * Direct function to sync student data between Group Lesson Tracker and SwimmerSkills sheets
 * This is a global function called directly from the menu
 */
function directSyncStudentData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    
    // Check if we're on the Group Lesson Tracker sheet
    if (sheet.getName() !== 'Group Lesson Tracker') {
      const ui = SpreadsheetApp.getUi();
      const result = ui.alert(
        'Switch to Group Lesson Tracker?',
        'This function should be run from the Group Lesson Tracker sheet. Would you like to switch to that sheet now?',
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        // Try to find and activate the Group Lesson Tracker sheet
        const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
        if (trackerSheet) {
          trackerSheet.activate();
          // Call sync function directly with the tracker sheet
          if (typeof GlobalFunctions !== 'undefined' && 
              typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
            GlobalFunctions.syncStudentDataWithSwimmerSkills(trackerSheet);
          } else {
            ui.alert('Error', 'Sync function not available. Please contact the administrator.', ui.ButtonSet.OK);
          }
        } else {
          ui.alert('Error', 'Could not find the Group Lesson Tracker sheet. Please create it first.', ui.ButtonSet.OK);
        }
      }
    } else {
      // We're already on the Group Lesson Tracker sheet
      if (typeof GlobalFunctions !== 'undefined' && 
          typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
        GlobalFunctions.syncStudentDataWithSwimmerSkills(sheet);
      } else {
        SpreadsheetApp.getUi().alert('Error', 'Sync function not available. Please contact the administrator.', SpreadsheetApp.getUi().ButtonSet.OK);
      }
    }
  } catch (error) {
    Logger.log(`Error in directSyncStudentData: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to sync student data: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Emergency function to fix menu issues
 * Redirects to the new central menu system's fixMenuSystem function
 */
function fixMenu() {
  // Call the new central menu system's fixMenuSystem function if available
  if (typeof fixMenuSystem === 'function') {
    return fixMenuSystem();
  }
  
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Force properties to true
    PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    // Call the new onOpen function from 00_MenuSystem.ts
    onOpen();
    
    // Show confirmation
    ui.alert(
      'Menu Fixed',
      'System properties have been reset and all menus have been created.',
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
 * Force recreate both main and Sync menus
 * This can be used to restore menus after an update
 * Now redirects to the new central menu system
 */
function reloadAllMenus() {
  // Call the new central menu system's functions if available
  if (typeof createMainMenu === 'function' && typeof createUtilityMenus === 'function') {
    try {
      createMainMenu();
      createUtilityMenus();
      
      SpreadsheetApp.getUi().alert(
        'Menus Reloaded',
        'All menus have been recreated using the new menu system.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
      return true;
    } catch (error) {
      Logger.log('Error using new menu system: ' + error.message);
    }
  }
  
  // Fallback to old method
  try {
    // Call the new onOpen function from 00_MenuSystem.ts
    onOpen();
    
    SpreadsheetApp.getUi().alert(
      'Menus Reloaded',
      'All menus have been recreated. Please check that both the YSL v6 Hub and utility menus appear in the menu bar.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    Logger.log('Error reloading menus: ' + error.message);
    SpreadsheetApp.getUi().alert('Error: ' + error.message);
    return false;
  }
}

/**
 * Special function to install an onOpen trigger
 * This function can be run from the Script Editor to install a trigger
 * that will run the onOpen function when the spreadsheet is opened
 * Now redirects to the new central menu system's fixTriggers function
 */
function installOnOpenTrigger() {
  // Call the new central menu system's fixTriggers function if available
  if (typeof fixTriggers === 'function') {
    return fixTriggers();
  }
  
  try {
    // Delete any existing onOpen triggers to avoid duplication
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === 'onOpen') {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
    
    // Create a new trigger for onOpen
    ScriptApp.newTrigger('onOpen')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
    
    Logger.log('Successfully installed onOpen trigger');
    return true;
  } catch (error) {
    Logger.log('Error installing onOpen trigger: ' + error.message);
    return false;
  }
}

/**
 * Function to create an example menu item that rebuilds menus
 * This can be run safely from the Script Editor
 */
function createSimpleRepairMenu() {
  try {
    // This function can be run from the Script Editor
    const ui = SpreadsheetApp.getUi();
    
    // Create a very minimal menu with just one repair function
    ui.createMenu('Menu Repair')
      .addItem('Fix Menu System', 'fixMenuSystem') // Use the new function name
      .addToUi();
    
    Logger.log('Created simple repair menu');
    return true;
  } catch (error) {
    Logger.log('Error creating simple repair menu: ' + error.message);
    return false;
  }
}

/**
 * Creates the full operational menu directly
 * This is now deprecated - the menu creation is handled by 00_MenuSystem.ts
 * This remains here for backward compatibility
 */
function createFullMenu() {
  // Call the new createMainMenu function from 00_MenuSystem.ts if available
  if (typeof createMainMenu === 'function') {
    return createMainMenu();
  }
  
  // Fallback to old implementation
  const ui = SpreadsheetApp.getUi();
  
  const menu = ui.createMenu('YSL v6 Hub');
  
  // 1. Class Management section - added directly to main menu instead of submenu
  menu.addItem('Generate Group Lesson Tracker', 'DynamicInstructorSheet_createDynamicInstructorSheet')
      .addItem('◉ SYNC STUDENT DATA ◉', 'directSyncStudentData') // Updated function reference
      .addSeparator()
      .addItem('Refresh Class List', 'DataIntegrationModule_updateClassSelector')
      .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData')
      .addItem('Generate Instructor Sheets', 'InstructorResourceModule_generateInstructorSheets')
      .addSeparator();
    
  // 2. Communications section
  menu.addSubMenu(ui.createMenu('Communications')
    .addItem('Create Communications Hub', 'CommunicationModule_createCommunicationsHub')
    .addItem('Create Communication Log', 'CommunicationModule_createCommunicationLog')
    .addItem('Send Selected Communication', 'CommunicationModule_sendSelectedCommunication')
    .addSeparator()
    .addItem('Send Mid-Session Reports', 'ReportingModule_generateMidSessionReports')
    .addItem('Send End-Session Reports', 'ReportingModule_generateEndSessionReports')
    .addSeparator()
    .addItem('Send Welcome Emails', 'CommunicationModule_sendWelcomeEmails')
    .addItem('Test Welcome Email', 'CommunicationModule_testWelcomeEmail'));
    
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
    .addItem('Fix Swimmer Records Access', 'fixSwimmerRecordsAccess')
    .addItem('Apply Configuration Changes', 'AdministrativeModule_applyConfigurationChanges')
    .addSeparator()
    .addItem('Show Logs', 'ErrorHandling_showLogViewer')
    .addItem('Show Version Info', 'VersionControl_showVersionInfo'));
    
  // 4. Tools & Diagnostics submenu
  menu.addSubMenu(ui.createMenu('Tools & Diagnostics')
    .addItem('System Health Check', 'DebugModule_performSystemHealthCheck')
    .addItem('System Diagnostics', 'VersionControl_showDiagnostics')
    .addSeparator()
    .addItem('Repair System', 'DebugModule_repairSystem')
    .addItem('Fix Menu', 'fixMenuSystem') // Updated function reference
    .addItem('Reload All Menus', 'reloadAllMenus')
    .addItem('Install Trigger', 'fixTriggers') // Updated function reference
    .addSeparator()
    .addItem('Clear System Cache', 'VersionControl_clearCache')
    .addItem('Test Menu Creation', 'testMenuCreation')); // Updated function reference
    
  // Add About item and the menu to UI
  menu.addSeparator()
    .addItem('About YSL v6 Hub', 'AdministrativeModule_showAboutDialog')
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
export function triggerOnEdit(e) {
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

/**
 * Test function for the sync functionality
 * This is a standalone function that can be called directly from the menu
 * It delegates to the GlobalFunctions.testSyncFunctionality method
 */
export function triggerTestSyncFunctionality() {
  try {
    // If GlobalFunctions is available, use its testSyncFunctionality function
    if (typeof GlobalFunctions !== 'undefined' && 
        typeof GlobalFunctions.testSyncFunctionality === 'function') {
      return GlobalFunctions.testSyncFunctionality();
    }
    
    // Fallback implementation if GlobalFunctions is not available
    SpreadsheetApp.getUi().alert(
      'Test Function Not Available',
      'The test function is not available. Please make sure the system is properly initialized.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (error) {
    Logger.log(`Error in testSyncFunctionality: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'Test Error',
      `An error occurred while testing the sync functionality: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}