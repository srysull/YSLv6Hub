/**
 * Global onOpen function that Google Apps Script will automatically recognize
 * and execute when the spreadsheet is opened.
 * 
 * This is a standalone function not part of any object, which ensures
 * Google Apps Script properly triggers it as an event handler.
 */
function onOpen() {
  try {
    // Call the main onOpen function if available
    if (typeof GlobalFunctions !== 'undefined' && typeof GlobalFunctions.onOpen === 'function') {
      GlobalFunctions.onOpen();
      return;
    }
    
    // Fallback to direct menu creation if GlobalFunctions isn't available
    Logger.log("Creating menu directly from standalone onOpen");
    
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('YSL Hub');
    
    try {
      // Try to use AdministrativeModule if available
      if (typeof AdministrativeModule !== 'undefined' && typeof AdministrativeModule.createMenu === 'function') {
        AdministrativeModule.createMenu();
        return;
      }
      
      // Direct fallback menu if AdministrativeModule isn't available
      menu.addItem('Initialize System', 'AdministrativeModule_showInitializationDialog')
          .addSeparator()
          .addItem('Repair System', 'AdministrativeModule_showConfigurationDialog')
          .addItem('Fix Menu', 'fixMenu')
          .addToUi();
    } catch (menuError) {
      Logger.log(`Menu creation error: ${menuError.message}`);
      
      // Last resort minimal menu
      ui.createMenu('YSL Hub')
        .addItem('Fix System', 'fixMenu')
        .addToUi();
    }
  } catch (error) {
    Logger.log(`Error in standalone onOpen: ${error.message}`);
    
    // Final attempt at a menu
    try {
      SpreadsheetApp.getUi()
        .createMenu('YSL Hub')
        .addItem('Emergency Fix', 'fixMenu')
        .addToUi();
    } catch (finalError) {
      Logger.log(`Final menu creation failed: ${finalError.message}`);
    }
  }
}

/**
 * Emergency function to fix menu issues
 */
function fixMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Set initialization properties
    PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    ui.alert(
      'Menu Repair',
      'System initialization properties have been set. Please reload the page to see the full menu.',
      ui.ButtonSet.OK
    );
    
    // Try to create the menu again
    if (typeof AdministrativeModule !== 'undefined' && typeof AdministrativeModule.createMenu === 'function') {
      AdministrativeModule.createMenu();
    } else {
      // Direct menu creation
      ui.createMenu('YSL Hub')
        .addItem('Initialize System', 'AdministrativeModule_showInitializationDialog')
        .addSeparator()
        .addItem('System Configuration', 'AdministrativeModule_showConfigurationDialog')
        .addToUi();
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error fixing menu: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to fix menu: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}