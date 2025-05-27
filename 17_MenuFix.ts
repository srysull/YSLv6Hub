/**
 * MenuFix.gs
 * 
 * This standalone script consolidates all menu-related fixes in one place.
 * Run these functions directly from the script editor to fix menu issues.
 */

/**
 * Primary menu fix function - run this first
 */
function completeMenuFix() {
  // Step 1: Fix properties
  PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
  PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
  
  // Step 2: Fix the trigger
  fixTriggers();
  
  // Step 3: Create the menu
  createFixedMenu();
  
  // Show confirmation
  SpreadsheetApp.getUi().alert(
    'Complete Menu Fix Applied',
    'The menu system has been fully repaired. You should now see the menu in the toolbar.\n\n' +
    'If you still do not see the menu after refreshing the page, please re-open the spreadsheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Fixes all triggers by removing duplicates and creating a single correct onOpen trigger
 * @deprecated Use fixTriggers from 00_MenuSystem.ts instead
 */
function fixTriggers_MenuFix() {
  // Delete any existing onOpen triggers to avoid duplication/conflict
  const triggers = ScriptApp.getProjectTriggers();
  let deletedCount = 0;
  
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(triggers[i]);
      deletedCount++;
    }
  }
  
  // Create a new trigger that points to the correct onOpen function in 00_TriggerFunctions.ts
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  Logger.log(`Fixed triggers: Deleted ${deletedCount} existing triggers and created a new onOpen trigger`);
  return true;
}

/**
 * Creates a comprehensive menu directly without relying on other functions
 * This serves as a failsafe menu creation method
 */
function createFixedMenu() {
  // Ensure properties are set before creating menus
  PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
  PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
  
  // Create the main menu
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('YSL v6 Hub');
  
  // 1. Class Management section
  menu.addItem('Generate Group Lesson Tracker', 'DynamicInstructorSheet_createDynamicInstructorSheet')
      .addItem('◉ SYNC STUDENT DATA ◉', 'directSyncStudentData')
      .addSeparator()
      .addItem('Refresh Class List', 'DataIntegrationModule_updateClassSelector')
      .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData')
      .addItem('Generate Instructor Sheets', 'InstructorResourceModule_generateInstructorSheets')
      .addSeparator();
    
  // 2. Communications section
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
    
  // 3. System section
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
    
  // 4. Tools & Diagnostics submenu
  menu.addSubMenu(ui.createMenu('Tools & Diagnostics')
    .addItem('System Health Check', 'DebugModule_performSystemHealthCheck')
    .addItem('System Diagnostics', 'VersionControl_showDiagnostics')
    .addSeparator()
    .addItem('Repair System', 'DebugModule_repairSystem')
    .addItem('Fix Menu', 'completeMenuFix')
    .addItem('Install Trigger', 'fixTriggers')
    .addSeparator()
    .addItem('Clear System Cache', 'VersionControl_clearCache')
    .addItem('Test Menu Creation', 'DebugModule_testMenuCreation'));
    
  // 5. Add Emergency Menu items directly to the main menu
  menu.addSeparator()
    .addItem('Fix Menu', 'completeMenuFix')
    .addItem('Fix Triggers', 'fixTriggers')
    .addSeparator()
    .addItem('About YSL v6 Hub', 'AdministrativeModule_showAboutDialog')
    .addToUi();
  
  // Create a separate sync menu
  ui.createMenu('Sync')
    .addItem('Sync Student Data', 'directSyncStudentData')
    .addToUi();
    
  // Also create a minimal emergency menu
  ui.createMenu('Emergency Repair')
    .addItem('Fix Menu System', 'completeMenuFix')
    .addItem('Fix Triggers', 'fixTriggers')
    .addItem('Run Menu Diagnostics', 'runMenuDiagnostics')
    .addToUi();
  
  Logger.log('Fixed menu created successfully');
  return true;
}

/**
 * Run diagnostics on the menu system and displays results
 * @deprecated Use runMenuDiagnostics from 00_MenuSystem.ts instead
 */
function runMenuDiagnostics_MenuFix() {
  // Check if key functions exist
  const diagnostics = {
    'createFullMenu': typeof createFullMenu === 'function',
    'directSyncStudentData': typeof directSyncStudentData === 'function',
    'fixMenu': typeof fixMenu === 'function',
    'installOnOpenTrigger': typeof installOnOpenTrigger === 'function',
    'reloadAllMenus': typeof reloadAllMenus === 'function',
    'onOpen in 00_TriggerFunctions': typeof onOpen === 'function',
    'createMenu in AdministrativeModule': typeof AdministrativeModule !== 'undefined' && 
                                         typeof AdministrativeModule.createMenu === 'function'
  };
  
  // Check initialization properties
  const props = PropertiesService.getScriptProperties();
  const initProps = {
    'systemInitialized': props.getProperty('systemInitialized'),
    'INITIALIZED': props.getProperty('INITIALIZED')
  };
  
  // Check triggers
  const triggers = ScriptApp.getProjectTriggers();
  const triggerInfo = triggers.map(trigger => ({
    'handlerFunction': trigger.getHandlerFunction(),
    'eventType': trigger.getEventType().toString(),
    'triggerSource': trigger.getTriggerSource().toString()
  }));
  
  // Display results
  let message = 'Menu System Diagnostics:\n\n';
  
  // Function checks
  message += 'Functions:\n';
  for (const [func, exists] of Object.entries(diagnostics)) {
    message += `- ${func}: ${exists ? 'EXISTS' : 'MISSING'}\n`;
  }
  
  // Property checks
  message += '\nInitialization Properties:\n';
  for (const [prop, value] of Object.entries(initProps)) {
    message += `- ${prop}: ${value || 'NOT SET'}\n`;
  }
  
  // Trigger checks
  message += '\nInstalled Triggers:\n';
  if (triggerInfo.length === 0) {
    message += '- NO TRIGGERS INSTALLED\n';
  } else {
    for (const trigger of triggerInfo) {
      message += `- ${trigger.handlerFunction} (${trigger.eventType})\n`;
    }
  }
  
  // Menu creation test
  message += '\nMenu Creation Test:\n';
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Test Menu')
      .addItem('Test Item', 'completeMenuFix')
      .addToUi();
    message += '- Menu creation test: SUCCESS\n';
  } catch (error) {
    message += `- Menu creation test: FAILED (${error.message})\n`;
  }
  
  // Recommendations
  message += '\nRecommendations:\n';
  
  if (triggerInfo.length === 0 || !triggerInfo.some(t => t.handlerFunction === 'onOpen')) {
    message += '- Run fixTriggers() to install the onOpen trigger correctly.\n';
  }
  
  if (!initProps.systemInitialized || !initProps.INITIALIZED) {
    message += '- Initialization properties are not set. Run completeMenuFix() to fix.\n';
  }
  
  // Multiple onOpen functions conflict
  message += '- There are two onOpen functions (in 00_TriggerFunctions.ts and 01_Globals.ts) which may cause conflicts.\n';
  message += '  Run fixTriggers() to ensure the trigger points to the correct function in 00_TriggerFunctions.ts.\n';
  
  // Show diagnostics
  SpreadsheetApp.getUi().alert('Menu Diagnostics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Force-refreshes the page by creating a temporary dialogue
 * This can help when changes don't appear immediately
 */
function forceRefresh() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Refresh Required',
    'The menu system has been fixed, but you need to refresh the page to see the changes.\n\n' +
    'Click OK and then refresh your browser.',
    ui.ButtonSet.OK
  );
}