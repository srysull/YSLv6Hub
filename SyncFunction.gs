/**
 * Synchronizes student data between Group Lesson Tracker and SwimmerSkills sheets.
 * This function is declared at the global scope to ensure it's accessible from the menu.
 */
function syncStudentData() {
  try {
    console.log("syncStudentData function called");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Find the Group Lesson Tracker sheet
    let trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    
    // If we don't have the sheet, alert the user and exit
    if (!trackerSheet) {
      ui.alert(
        'Error',
        'Group Lesson Tracker sheet not found. Please create it first.',
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Get the active sheet to check if we're already on Group Lesson Tracker
    const currentSheet = ss.getActiveSheet();
    
    // If we're not on the Group Lesson Tracker sheet, ask to switch
    if (currentSheet.getName() !== 'Group Lesson Tracker') {
      const result = ui.alert(
        'Switch to Group Lesson Tracker?',
        'This function should be run from the Group Lesson Tracker sheet. Would you like to switch to that sheet now?',
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        // Activate the Group Lesson Tracker sheet
        trackerSheet.activate();
      } else {
        // User chose not to switch sheets
        return;
      }
    }
    
    // Now run the synchronization
    // We'll use direct access to the sync function in Globals
    if (typeof GlobalFunctions !== 'undefined' && 
        typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
      
      // Call the actual sync function with the tracker sheet
      GlobalFunctions.syncStudentDataWithSwimmerSkills(trackerSheet);
      
      // Show success
      ui.alert(
        'Sync Complete',
        'Student data has been synchronized between Group Lesson Tracker and SwimmerSkills.',
        ui.ButtonSet.OK
      );
    } else {
      // Fallback if GlobalFunctions is not available
      ui.alert(
        'Error',
        'Sync function not available. Try refreshing the page or contact support.',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    console.error("Error in syncStudentData: " + error.message);
    
    // Alert the user of the error
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to sync student data: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}