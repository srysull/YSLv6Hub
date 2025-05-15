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
    
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('YSL Hub')
      .addItem('Initialize System', 'AdministrativeModule_showInitializationDialog')
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
    // Force properties to true
    PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    // Create the full menu
    createFullMenu();
    
    // Show confirmation
    SpreadsheetApp.getUi().alert(
      'Menu Fixed',
      'System properties have been reset and menu has been created.',
      SpreadsheetApp.getUi().ButtonSet.OK
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
  var ui = SpreadsheetApp.getUi();
  
  var menu = ui.createMenu('YSL Hub');
  
  // 1. Class Management section
  menu.addSubMenu(ui.createMenu('Class Management')
    .addItem('Create Dynamic Class Hub', 'DynamicInstructorSheet_createDynamicInstructorSheet')
    .addItem('Update with Selected Class', 'DynamicInstructorSheet_rebuildDynamicInstructorSheet')
    .addSeparator()
    .addItem('Refresh Class List', 'DataIntegrationModule_updateClassSelector')
    .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData'));
    
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
    .addItem('Show Logs', 'ErrorHandling_showLogViewer'));
    
  // Add About item and the menu to UI
  menu.addSeparator()
    .addItem('Fix Menu', 'fixMenu')
    .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog')
    .addToUi();
    
  return menu;
}

/**
 * Global onEdit trigger function
 * Executed automatically when the spreadsheet is edited
 * 
 * @param {Object} e - The edit event object
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
  } catch (error) {
    Logger.log(`Error in onEdit trigger: ${error.message}`);
  }
}