/**
 * Trigger Installation Script
 * 
 * This script ensures the onOpen trigger is properly installed
 * Run this manually if the menu stops appearing
 */

/**
 * Install the onOpen trigger
 * This function can be run manually from the script editor
 */
function installOnOpenTrigger() {
  try {
    // First, remove any existing onOpen triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    let removedCount = 0;
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onOpen') {
        ScriptApp.deleteTrigger(trigger);
        removedCount++;
        Logger.log(`Removed existing onOpen trigger for function: ${trigger.getHandlerFunction()}`);
      }
    }
    
    // Create new onOpen trigger
    const newTrigger = ScriptApp.newTrigger('onOpen')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
    
    Logger.log(`Trigger installation complete. Removed ${removedCount} old triggers, created 1 new trigger.`);
    Logger.log(`New trigger ID: ${newTrigger.getUniqueId()}`);
    
    // Test the trigger by calling onOpen
    Logger.log('Testing onOpen trigger...');
    onOpen();
    
    // Show success message
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Trigger Installed',
      `Successfully installed onOpen trigger. Removed ${removedCount} old triggers and created 1 new trigger. The menu should now appear when you reload the spreadsheet.`,
      ui.ButtonSet.OK
    );
    
    return true;
    
  } catch (error) {
    Logger.log(`Error installing onOpen trigger: ${error.message}`);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Installation Failed',
      `Failed to install onOpen trigger: ${error.message}`,
      ui.ButtonSet.OK
    );
    
    return false;
  }
}

/**
 * Check the status of onOpen triggers
 */
function checkOnOpenTriggers() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const onOpenTriggers = triggers.filter(t => t.getHandlerFunction() === 'onOpen');
    
    let message = `Found ${onOpenTriggers.length} onOpen trigger(s):\n\n`;
    
    if (onOpenTriggers.length === 0) {
      message += 'No onOpen triggers found! This is why the menu is not appearing.\n\n';
      message += 'Click OK, then run "installOnOpenTrigger" to fix this.';
    } else {
      onOpenTriggers.forEach((trigger, index) => {
        message += `${index + 1}. Function: ${trigger.getHandlerFunction()}\n`;
        message += `   Trigger ID: ${trigger.getUniqueId()}\n`;
        message += `   Event Type: ${trigger.getEventType()}\n\n`;
      });
      
      if (onOpenTriggers.length > 1) {
        message += 'Warning: Multiple onOpen triggers detected. This might cause issues.\n';
        message += 'Consider running "installOnOpenTrigger" to clean up duplicates.';
      }
    }
    
    const ui = SpreadsheetApp.getUi();
    ui.alert('OnOpen Trigger Status', message, ui.ButtonSet.OK);
    
    // Also log the information
    Logger.log(`OnOpen Trigger Check: ${onOpenTriggers.length} triggers found`);
    onOpenTriggers.forEach(trigger => {
      Logger.log(`- ${trigger.getHandlerFunction()} (${trigger.getUniqueId()})`);
    });
    
    return onOpenTriggers.length;
    
  } catch (error) {
    Logger.log(`Error checking onOpen triggers: ${error.message}`);
    
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Check Failed',
      `Failed to check onOpen triggers: ${error.message}`,
      ui.ButtonSet.OK
    );
    
    return -1;
  }
}