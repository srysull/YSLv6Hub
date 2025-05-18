/**
 * InstallTrigger.gs
 * 
 * This standalone script ensures the onOpen trigger is properly installed.
 * It should be run manually from the script editor.
 */

/**
 * Run this function from the script editor to install the onOpen trigger
 */
function manualInstallTrigger() {
  // Delete any existing onOpen triggers to avoid duplication
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onOpen') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // Create a new trigger for onOpen
  ScriptApp.newTrigger('onOpen')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onOpen()
    .create();
  
  // Fix initialization properties
  PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
  PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
  
  // Create a simple repair menu immediately
  try {
    // Call the improved menu fix function if available
    if (typeof completeMenuFix === 'function') {
      completeMenuFix();
    } else {
      // Fallback to basic menu creation
      const ui = SpreadsheetApp.getUi();
      ui.createMenu('Emergency Repair')
        .addItem('Fix Menu', 'fixMenu')
        .addItem('Install Trigger', 'manualInstallTrigger')
        .addToUi();
    }
  } catch (menuError) {
    Logger.log('Error creating menu: ' + menuError.message);
  }
  
  // Show confirmation
  SpreadsheetApp.getUi().alert(
    'Trigger Installed',
    'The onOpen trigger has been successfully installed. ' +
    'Please refresh or reopen the spreadsheet to see the menu.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Run this function to regenerate the menu system manually
 */
function manualCreateMenus() {
  // Force properties to true to ensure menu appears
  PropertiesService.getScriptProperties().setProperty('systemInitialized', 'true');
  PropertiesService.getScriptProperties().setProperty('INITIALIZED', 'true');
  
  // Call the improved menu fix function if available
  if (typeof completeMenuFix === 'function') {
    completeMenuFix();
    return;
  }
  
  // Create the full menu if it exists
  if (typeof createFullMenu === 'function') {
    createFullMenu();
  } else {
    // Create a minimal menu
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('YSL v6 Hub')
      .addItem('Fix Menu', 'fixMenu')
      .addItem('Install Trigger', 'manualInstallTrigger')
      .addItem('Reload Menus', 'reloadAllMenus')
      .addToUi();
  }
  
  // Create sync menu if possible
  if (typeof createSyncMenu === 'function') {
    createSyncMenu();
  } else {
    // Create a minimal sync menu
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('Sync')
      .addItem('Sync Student Data', 'directSyncStudentData')
      .addToUi();
  }
  
  // Show confirmation
  SpreadsheetApp.getUi().alert(
    'Menus Created',
    'Menus have been manually created. You should now see them in the menu bar.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Run this function to diagnose menu issues
 */
function menuDiagnostics() {
  // Forward to the improved diagnostics function if available
  if (typeof runMenuDiagnostics === 'function') {
    runMenuDiagnostics();
    return;
  }
  
  // Check if key functions exist
  const diagnostics = {
    'createFullMenu': typeof createFullMenu === 'function',
    'directSyncStudentData': typeof directSyncStudentData === 'function',
    'fixMenu': typeof fixMenu === 'function',
    'installOnOpenTrigger': typeof installOnOpenTrigger === 'function',
    'reloadAllMenus': typeof reloadAllMenus === 'function',
    'onOpen': typeof onOpen === 'function'
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
  let message = 'Menu Diagnostics:\n\n';
  
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
  
  // Recommendations
  message += '\nRecommendations:\n';
  if (!diagnostics.onOpen) {
    message += '- The onOpen function is missing. This is critical for menu creation.\n';
  }
  if (triggerInfo.length === 0 || !triggerInfo.some(t => t.handlerFunction === 'onOpen')) {
    message += '- Run manualInstallTrigger() to install the onOpen trigger.\n';
  }
  if (!initProps.systemInitialized && !initProps.INITIALIZED) {
    message += '- Initialization properties are not set. Run manualCreateMenus() to fix.\n';
  }
  
  // Show diagnostics
  SpreadsheetApp.getUi().alert('Menu Diagnostics', message, SpreadsheetApp.getUi().ButtonSet.OK);
}