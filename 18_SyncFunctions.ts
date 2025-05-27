/**
 * YSL Hub v2 Data Synchronization Functions
 * 
 * This module provides functions for synchronizing data between different sheets.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

/**
 * Synchronizes student data between Group Lesson Tracker and SwimmerSkills sheets
 * This global function is designed to be called from the menu
 */
function syncSwimmerData() {
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
          // Call the sync function with the correct sheet
          if (typeof GlobalFunctions !== 'undefined' && 
              typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
            GlobalFunctions.syncStudentDataWithSwimmerSkills(trackerSheet);
          } else {
            // Direct fallback if GlobalFunctions is not available
            // This should not happen, but log error if it does
            Logger.log('ERROR: GlobalFunctions.syncStudentDataWithSwimmerSkills is not available');
          }
          return true;
        } else {
          ui.alert('Error', 'Could not find the Group Lesson Tracker sheet. Please create it first.', ui.ButtonSet.OK);
          return false;
        }
      } else {
        // User chose not to switch sheets
        return false;
      }
    }
    
    // Call the sync function directly
    if (typeof GlobalFunctions !== 'undefined' && 
        typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
      GlobalFunctions.syncStudentDataWithSwimmerSkills(sheet);
    } else {
      // Direct fallback if GlobalFunctions is not available
      // This should not happen, but log error if it does
      Logger.log('ERROR: GlobalFunctions.syncStudentDataWithSwimmerSkills is not available');
    }
    
    return true;
  } catch (error) {
    Logger.log(`Error in syncSwimmerData: ${error.message}`);
    SpreadsheetApp.getUi().alert('Error', `Failed to sync student data: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
}