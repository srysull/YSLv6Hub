/**
 * YSL_SYNC_STUDENT_DATA
 * 
 * Global function to sync student data between Group Lesson Tracker and SwimmerSkills
 * This function is called directly from the YSL v6 Hub menu.
 */
function YSL_SYNC_STUDENT_DATA() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    
    // Find the necessary sheets
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    const skillsSheet = ss.getSheetByName('SwimmerSkills');
    
    // Validate sheets exist
    if (!trackerSheet) {
      ui.alert('Error', 'Group Lesson Tracker sheet not found. Please create it first.', ui.ButtonSet.OK);
      return;
    }
    
    if (!skillsSheet) {
      ui.alert('Error', 'SwimmerSkills sheet not found. Please create it first.', ui.ButtonSet.OK);
      return;
    }
    
    // If not on Group Lesson Tracker, offer to switch
    if (ss.getActiveSheet().getName() !== 'Group Lesson Tracker') {
      const result = ui.alert(
        'Switch to Group Lesson Tracker?',
        'This function should be run from the Group Lesson Tracker sheet. Would you like to switch to that sheet now?',
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        trackerSheet.activate();
      } else {
        return;
      }
    }
    
    // Call the sync function if available
    if (typeof GlobalFunctions !== 'undefined' && 
        typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
      
      GlobalFunctions.syncStudentDataWithSwimmerSkills(trackerSheet);
      
      ui.alert(
        'Sync Complete',
        'Student data has been synchronized between Group Lesson Tracker and SwimmerSkills.',
        ui.ButtonSet.OK
      );
    } else {
      // Implement a direct sync if GlobalFunctions is not available
      ui.alert(
        'Error',
        'The sync function is not available. Please contact the administrator.',
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    // Handle any errors
    SpreadsheetApp.getUi().alert(
      'Error',
      'Failed to sync student data: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}