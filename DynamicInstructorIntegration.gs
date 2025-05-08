/**
 * YSL Hub Dynamic Instructor Integration Module
 * 
 * This module provides integration points for the Dynamic Instructor Sheet
 * with the rest of the YSL Hub system, including menu items and initialization.
 * 
 * @author Claude Code
 * @version 1.0
 * @date 2025-05-05
 */

/**
 * Adds Dynamic Instructor Sheet menu items to the main menu
 * @param {Menu} menu - The menu to add items to
 * @return {Menu} The updated menu
 */
function addDynamicInstructorMenuItems(menu) {
  return menu.addItem('Create Dynamic Instructor Sheet', 'DynamicInstructorSheet.createDynamicInstructorSheet')
             .addItem('Push Skills to Swimmer Records', 'DynamicInstructorSheet.pushSkillsToSwimmerRecords');
}

/**
 * Updates the YSL Hub menu with Dynamic Instructor Sheet options
 */
function updateMenuWithDynamicInstructor() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('YSL Hub');
    
    // Check if system is initialized
    const isInitialized = GlobalFunctions.safeGetProperty('systemInitialized') === 'true';
    
    if (isInitialized) {
      menu.addSubMenu(ui.createMenu('Class Management')
            .addItem('Update Class Selector', 'DataIntegrationModule_updateClassSelector')
            .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData')
            .addItem('Update Instructor Data', 'DataIntegrationModule_updateInstructorData')
            .addItem('Create Dynamic Instructor Sheet', 'DynamicInstructorSheet.createDynamicInstructorSheet')
            .addItem('Push Skills to Swimmer Records', 'DynamicInstructorSheet.pushSkillsToSwimmerRecords'));
    } else {
      menu.addItem('Initialize System', 'AdministrativeModule.showInitializationDialog');
    }
    
    menu.addSeparator()
        .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog')
        .addToUi();
    
    return menu;
  } catch (error) {
    Logger.log(`Error updating menu with dynamic instructor: ${error.message}`);
    return null;
  }
}

/**
 * Initialize the dynamic instructor module
 * This is called during system initialization to set up necessary components
 * @return {boolean} Success status
 */
function initializeDynamicInstructorModule() {
  try {
    // No initialization steps needed yet - the sheet is created on demand
    return true;
  } catch (error) {
    Logger.log(`Error initializing dynamic instructor module: ${error.message}`);
    return false;
  }
}

/**
 * Update the original InstructorResourceModule to integrate with the dynamic sheet
 */
function updateInstructorResourceModule() {
  try {
    // Nothing to do - we're not modifying the original module
    // The dynamic sheet exists alongside the original functionality
    return true;
  } catch (error) {
    Logger.log(`Error updating instructor resource module: ${error.message}`);
    return false;
  }
}

// Make functions available to other modules
const DynamicInstructorIntegration = {
  addDynamicInstructorMenuItems: addDynamicInstructorMenuItems,
  updateMenuWithDynamicInstructor: updateMenuWithDynamicInstructor,
  initializeDynamicInstructorModule: initializeDynamicInstructorModule
};