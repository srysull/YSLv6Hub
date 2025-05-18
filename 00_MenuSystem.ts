/**
 * YSL Hub v6 Menu System
 * Central menu implementation for the YSL Hub system
 * 
 * This file contains the core menu functionality for the YSL Hub application.
 * It replaces the scattered menu implementations to provide a single, reliable
 * menu system. By centralizing the menu code, we eliminate conflicts that
 * previously caused issues.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-18
 */

// Global menu configuration
const MENU_CONFIG = {
  title: 'YSL v6 Hub',
  properties: {
    initialized: 'systemInitialized',
    version: 'systemVersion'
  }
}

/**
 * Primary onOpen function - called when the spreadsheet is opened
 * This should be the ONLY onOpen function in the entire project
 */
function onOpen() {
  try {
    // Log the start of the function
    Logger.log('onOpen trigger function started from 00_MenuSystem.ts');
    
    // Always set initialization properties
    PropertiesService.getScriptProperties().setProperty(MENU_CONFIG.properties.initialized, 'true');
    PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
    
    // Create main menu only
    createMainMenu();
    
    Logger.log('Menu creation completed successfully');
  } catch (error) {
    Logger.log('Error in onOpen function: ' + error.message);
    try {
      // Emergency fallback menu if normal creation fails
      createEmergencyMenu();
    } catch (fallbackError) {
      Logger.log('Failed to create emergency menu: ' + fallbackError.message);
    }
  }
}

/**
 * Creates the main YSL Hub menu
 */
function createMainMenu() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(MENU_CONFIG.title);
  
  // Class Management
  menu.addItem('Generate Group Lesson Tracker', 'DynamicInstructorSheet_createDynamicInstructorSheet')
      .addItem('Sync Student Data', 'directSyncStudentData')
      .addSeparator()
      .addItem('Refresh Class List', 'DataIntegrationModule_updateClassSelector')
      .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData')
      .addItem('Generate Instructor Sheets', 'InstructorResourceModule_generateInstructorSheets')
      .addSeparator();
  
  // Communications submenu
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
  
  // System submenu
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
  
  // Tools & Diagnostics submenu
  menu.addSubMenu(ui.createMenu('Tools & Diagnostics')
    .addItem('System Health Check', 'DebugModule_performSystemHealthCheck')
    .addItem('System Diagnostics', 'VersionControl_showDiagnostics')
    .addSeparator()
    .addItem('Repair System', 'DebugModule_repairSystem')
    .addItem('Fix Menu', 'fixMenuSystem')
    .addItem('Run Menu Diagnostics', 'runMenuDiagnostics')
    .addSeparator()
    .addItem('Clear System Cache', 'VersionControl_clearCache')
    .addItem('Test Menu Creation', 'testMenuCreation'));
  
  // Final menu items
  menu.addSeparator()
    .addItem('About YSL v6 Hub', 'AdministrativeModule_showAboutDialog')
    .addToUi();
}

/**
 * Creates utility menus for quick access to important functions
 * Now empty - all functionality moved to main menu
 */
function createUtilityMenus() {
  // No additional menus - all functionality is in the main menu
  // This function is kept for backward compatibility
}

/**
 * Creates a minimal emergency menu when regular menu creation fails
 */
function createEmergencyMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Menu Repair')
    .addItem('Fix Menu System', 'fixMenuSystem')
    .addToUi();
}

/**
 * Fixes menu system issues
 */
function fixMenuSystem() {
  // Set initialization properties
  PropertiesService.getScriptProperties().setProperty(MENU_CONFIG.properties.initialized, 'true');
  PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
  
  // Fix triggers by ensuring only one onOpen trigger exists
  fixTriggers();
  
  // Create main menu only
  createMainMenu();
  
  // Show confirmation
  SpreadsheetApp.getUi().alert(
    'Menu System Fixed',
    'The menu system has been repaired. The menu should now be visible in the toolbar.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Fixes trigger issues by removing duplicates and ensuring single onOpen
 */
function fixTriggers() {
  // Delete existing onOpen triggers
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(trigger);
      deletedCount++;
    }
  }
  
  // Create a new trigger for onOpen
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  Logger.log(`Fixed triggers: Deleted ${deletedCount} existing triggers and created a new onOpen trigger`);
  return true;
}

/**
 * Tests menu creation
 */
function testMenuCreation() {
  // Test menu creation and log results
  try {
    createMainMenu();
    createUtilityMenus();
    Logger.log('Menu creation test successful');
    
    SpreadsheetApp.getUi().alert(
      'Menu Creation Test',
      'Menu creation test was successful. All menus should now be visible.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    Logger.log('Menu creation test failed: ' + error.message);
    
    SpreadsheetApp.getUi().alert(
      'Menu Creation Test Failed',
      `Menu creation test failed: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false;
  }
}

/**
 * Runs diagnostics on the menu system
 */
function runMenuDiagnostics() {
  // Properties check
  const props = PropertiesService.getScriptProperties();
  const initProps = {
    [MENU_CONFIG.properties.initialized]: props.getProperty(MENU_CONFIG.properties.initialized),
    'INITIALIZED': props.getProperty('INITIALIZED')
  };
  
  // Trigger check
  const triggers = ScriptApp.getProjectTriggers();
  const onOpenTriggers = triggers.filter(t => t.getHandlerFunction() === 'onOpen');
  
  // Display results
  let message = 'Menu System Diagnostics:\n\n';
  
  // Property checks
  message += 'Initialization Properties:\n';
  for (const [prop, value] of Object.entries(initProps)) {
    message += `- ${prop}: ${value || 'NOT SET'}\n`;
  }
  
  // Trigger checks
  message += '\nOnOpen Triggers:\n';
  if (onOpenTriggers.length === 0) {
    message += '- NO ONOPEN TRIGGERS INSTALLED\n';
  } else {
    message += `- Found ${onOpenTriggers.length} onOpen triggers\n`;
  }
  
  // Recommendations
  message += '\nRecommendations:\n';
  
  if (onOpenTriggers.length === 0) {
    message += '- Run fixTriggers() to install the onOpen trigger\n';
  } else if (onOpenTriggers.length > 1) {
    message += '- Run fixTriggers() to remove duplicate onOpen triggers\n';
  }
  
  if (!initProps[MENU_CONFIG.properties.initialized] || !initProps['INITIALIZED']) {
    message += '- Initialization properties are not set. Run fixMenuSystem() to fix.\n';
  }
  
  // Show diagnostics
  SpreadsheetApp.getUi().alert('Menu Diagnostics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}