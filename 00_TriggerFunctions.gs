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
    // Create a minimal menu immediately as a fallback
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('YSL Hub')
      .addItem('Initialize System', 'AdministrativeModule_showInitializationDialog')
      .addItem('System Configuration', 'AdministrativeModule_showConfigurationDialog')
      .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog')
      .addToUi();
    
    // Try using the more robust menu system
    if (typeof AdministrativeModule !== 'undefined' && 
        typeof AdministrativeModule.createMenu === 'function') {
      AdministrativeModule.createMenu();
    }
    
    // Initialize error handling if available
    if (typeof ErrorHandling !== 'undefined' && 
        typeof ErrorHandling.initializeErrorHandling === 'function') {
      ErrorHandling.initializeErrorHandling();
    }
    
    // Initialize version control if available
    if (typeof VersionControl !== 'undefined' && 
        typeof VersionControl.initializeVersionControl === 'function') {
      VersionControl.initializeVersionControl();
    }
    
    // Log the opening event
    Logger.log('Spreadsheet opened - onOpen trigger executed');
  } catch (error) {
    // Log error as a last resort
    Logger.log(`Error in onOpen trigger: ${error.message}`);
    
    // Create an emergency menu
    try {
      SpreadsheetApp.getUi()
        .createMenu('YSL Emergency')
        .addItem('Fix System', 'AdministrativeModule_fixSystemInitializationProperty')
        .addToUi();
    } catch (finalError) {
      Logger.log(`Failed to create emergency menu: ${finalError.message}`);
    }
  }
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