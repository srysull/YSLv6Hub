/**
 * YSL Hub Dynamic Instructor Integration Module
 * 
 * This module provides integration points for the Dynamic Instructor Sheet
 * with the rest of the YSL Hub system, including menu items and initialization.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

/**
 * Adds Dynamic Instructor Sheet menu items to the main menu
 * @param {Menu} menu - The menu to add items to
 * @return {Menu} The updated menu
 */
function addDynamicInstructorMenuItems(menu) {
  return menu
    .addItem('Create New Instructor Sheet', 'DynamicInstructorSheet.createDynamicInstructorSheet')
    .addSeparator()
    .addItem('Update Instructor Sheet with Selected Class', 'DynamicInstructorSheet.rebuildDynamicInstructorSheet');
}

/**
 * Updates the YSL Hub menu with Dynamic Instructor Sheet options
 */
// This function is no longer used since we've integrated the menu directly in AdministrativeModule.createMenu
function updateMenuWithDynamicInstructor() {
  try {
    Logger.log('Using the integrated menu structure in AdministrativeModule instead');
    return null;
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