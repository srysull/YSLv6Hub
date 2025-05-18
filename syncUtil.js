/**
 * SYNC UTILITY FUNCTIONS
 * Global simple functions for syncing data
 */

// Function to add the sync menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('↺ Sync')
    .addItem('Sync Student Data', 'syncData')
    .addToUi();
}

// Function to sync data
function syncData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Find tracker sheet
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    if (!trackerSheet) {
      ui.alert('Error', 'Group Lesson Tracker sheet not found. Please create it first.', ui.ButtonSet.OK);
      return;
    }
    
    // If not on the tracker sheet, offer to switch
    if (ss.getActiveSheet().getName() !== 'Group Lesson Tracker') {
      const response = ui.alert(
        'Switch Sheet?', 
        'This function should be run from the Group Lesson Tracker sheet. Switch now?', 
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        trackerSheet.activate();
      } else {
        return;
      }
    }
    
    // Do the sync
    if (typeof GlobalFunctions !== 'undefined' && 
        typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
      
      // Call the real sync function
      GlobalFunctions.syncStudentDataWithSwimmerSkills(trackerSheet);
      
      // Confirm success
      ui.alert('Sync Complete', 'Data has been synchronized successfully.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'Sync function not available.', ui.ButtonSet.OK);
    }
  } catch (error) {
    ui.alert('Error', 'Failed to sync data: ' + error.message, ui.ButtonSet.OK);
  }
}