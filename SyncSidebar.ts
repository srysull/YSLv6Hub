/**
 * Function to show the sync sidebar
 * This will be directly added to the menu
 */
function showSyncSidebar() {
  // Create the sidebar
  var html = HtmlService.createHtmlOutputFromFile('SyncUI')
      .setTitle('YSL Data Sync')
      .setWidth(300);
  
  // Display the sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Function called from the sidebar to run the sync operation
 * Returns a status message to display in the sidebar
 */
function runSyncOperation() {
  try {
    // Get the sheets
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    
    // If Group Lesson Tracker sheet doesn't exist, show error
    if (!trackerSheet) {
      return "Error: Group Lesson Tracker sheet not found. Create it first.";
    }
    
    // Activate the Group Lesson Tracker sheet
    trackerSheet.activate();
    
    // Call the sync function
    if (typeof GlobalFunctions !== 'undefined' && 
        typeof GlobalFunctions.syncStudentDataWithSwimmerSkills === 'function') {
      
      // Call the sync function from GlobalFunctions
      GlobalFunctions.syncStudentDataWithSwimmerSkills(trackerSheet);
      
      // Return success message
      return "Sync complete! Data has been synchronized between Group Lesson Tracker and SwimmerSkills sheets.";
    } else {
      // Sync function not available
      return "Error: Sync function not available. Please contact administrator.";
    }
  } catch (error) {
    // Handle any errors
    return "Error: " + error.message;
  }
}