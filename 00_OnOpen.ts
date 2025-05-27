/**
 * YSL Hub v2 OnOpen Trigger
 * 
 * Simple, standalone onOpen function to ensure menu always loads
 * This file exists solely to guarantee the onOpen trigger works
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-27
 */

/**
 * Main onOpen trigger function
 * Called automatically when the spreadsheet is opened
 */
function onOpen() {
  try {
    Logger.log('onOpen trigger started from 00_OnOpen.ts');
    
    // Set initialization properties
    PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    // Create the main menu
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('YSL v6 Hub');
    
    // Add main menu items
    menu.addItem('Generate Group Lesson Tracker', 'DynamicInstructorSheet_createDynamicInstructorSheet')
        .addItem('Sync Student Data', 'directSyncStudentData')
        .addSeparator()
        .addItem('Refresh Class List', 'DataIntegrationModule_updateClassSelector')
        .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData')
        .addItem('Generate Instructor Sheets', 'InstructorResourceModule_generateInstructorSheets')
        .addSeparator();
    
    // Add Communications submenu
    menu.addSubMenu(ui.createMenu('Communications')
        .addItem('Create Communications Hub', 'CommunicationModule_createCommunicationsHub')
        .addItem('Create Communication Log', 'CommunicationModule_createCommunicationLog')
        .addItem('Send Selected Communication', 'CommunicationModule_sendSelectedCommunication')
        .addSeparator()
        .addItem('Send Mid-Session Reports', 'ReportingModule_generateMidSessionReports')
        .addItem('Send End-Session Reports', 'ReportingModule_generateEndSessionReports')
        .addSeparator()
        .addItem('Send Welcome Emails', 'CommunicationModule_sendWelcomeEmails')
        .addItem('Test Welcome Email', 'CommunicationModule_testWelcomeEmail'));
    
    // Add System submenu
    menu.addSubMenu(ui.createMenu('System')
        .addItem('Create User Guide', 'UserGuide_createUserGuideSheet')
        .addSeparator()
        .addItem('View History', 'HistoryModule_createHistorySheet')
        .addSeparator()
        .addItem('Start New Session', 'SessionTransitionModule_startSessionTransition')
        .addItem('Resume Session Transition', 'SessionTransitionModule_resumeSessionTransition')
        .addSeparator()
        .addItem('System Configuration', 'AdministrativeModule_showConfigurationDialog')
        .addItem('Fix Swimmer Records Access', 'fixSwimmerRecordsAccess')
        .addItem('Apply Configuration Changes', 'AdministrativeModule_applyConfigurationChanges')
        .addSeparator()
        .addItem('Show Logs', 'ErrorHandling_showLogViewer')
        .addItem('Show Version Info', 'VersionControl_showVersionInfo'));
    
    // Add Tools & Diagnostics submenu
    menu.addSubMenu(ui.createMenu('Tools & Diagnostics')
        .addItem('System Health Check', 'DebugModule_performSystemHealthCheck')
        .addItem('System Diagnostics', 'VersionControl_showDiagnostics')
        .addSeparator()
        .addItem('Repair System', 'DebugModule_repairSystem')
        .addItem('Fix Menu', 'fixMenuSystem')
        .addItem('Run Menu Diagnostics', 'runMenuDiagnostics')
        .addSeparator()
        .addItem('Test Sync Functionality', 'testSyncFunctionality')
        .addSeparator()
        .addItem('Clear System Cache', 'VersionControl_clearCache')
        .addItem('Test Menu Creation', 'testMenuCreation'));
    
    // Add final items and show menu
    menu.addSeparator()
        .addItem('About YSL v6 Hub', 'AdministrativeModule_showAboutDialog')
        .addToUi();
    
    Logger.log('Menu creation completed successfully from 00_OnOpen.ts');
    
  } catch (error) {
    Logger.log('Error in onOpen from 00_OnOpen.ts: ' + error.message);
    
    // Emergency fallback - create simple menu
    try {
      const ui = SpreadsheetApp.getUi();
      ui.createMenu('YSL v6 Hub - Emergency')
        .addItem('Fix Menu System', 'fixMenuSystem')
        .addItem('Repair System', 'DebugModule_repairSystem')
        .addToUi();
      Logger.log('Emergency menu created');
    } catch (fallbackError) {
      Logger.log('Failed to create emergency menu: ' + fallbackError.message);
    }
  }
}

/**
 * Function to manually fix menu issues
 */
function fixMenuSystem() {
  try {
    // Set properties
    PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    // Remove existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onOpen') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }
    
    // Create new trigger
    ScriptApp.newTrigger('onOpen')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
    
    Logger.log(`Fixed triggers: Deleted ${deletedCount} existing triggers and created a new onOpen trigger`);
    
    // Call onOpen to recreate menu immediately
    onOpen();
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      'Menu System Fixed',
      'The menu system has been repaired. The menu should now be visible in the toolbar.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    Logger.log('Error in fixMenuSystem: ' + error.message);
    SpreadsheetApp.getUi().alert(
      'Fix Failed',
      'Failed to fix menu system: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}