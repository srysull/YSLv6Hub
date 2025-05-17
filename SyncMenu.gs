/**
 * Creates a standalone Sync menu
 * This should be called from the main onOpen function
 */
function createSyncMenu() {
  // Create a simple menu that directly calls the sync function
  SpreadsheetApp.getUi()
    .createMenu('Sync')
    .addItem('Sync Student Data', 'syncSwimmerData_menuWrapper')
    .addToUi();
}