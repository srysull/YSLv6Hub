/**
 * YSLv6Hub System Module
 * 
 * This module handles system initialization, menu creation, and dashboard functionality.
 * It serves as the entry point for the application and manages the overall system state.
 * 
 * @author Sean R. Sullivan
 * @version 1.0.0
 * @date 2025-05-18
 */

/**
 * Runs when the spreadsheet is opened
 * Sets up menus and initializes the system
 */
function onOpen(): void {
  try {
    // Log that we're starting
    console.log('Starting YSLv6Hub System initialization');
    
    // Set up the menu
    createMenu();
    
    // Initialize the system
    initializeSystem();
    
    console.log('YSLv6Hub System initialization complete');
  } catch (error) {
    console.error('Error in onOpen:', error);
    
    // Create emergency menu if regular initialization fails
    try {
      createEmergencyMenu();
    } catch (menuError) {
      console.error('Failed to create emergency menu:', String(menuError));
    }
  }
}

/**
 * Creates the main application menu
 */
function createMenu(): void {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('YSLv6Hub');
  
  // Basic menu items
  menu.addItem('Dashboard', 'showDashboard')
      .addSeparator();
  
  // Registration menu
  menu.addItem('Import Registration Data', 'importRegistrationData')
      .addItem('Manage Classes', 'manageClasses')
      .addSeparator();
  
  // Tracking menu
  menu.addItem('Generate GroupsTracker', 'generateGroupsTracker')
      .addItem('Sync Data', 'syncData')
      .addSeparator();
  
  // Communications menu
  menu.addSubMenu(ui.createMenu('Communications')
      .addItem('Communications Hub', 'showCommsHub')
      .addItem('Generate Reports', 'generateReports')
      .addItem('Group Communications', 'showGroupComms')
      .addItem('Instructor Communications', 'showInstructorComms'));
  
  // Session management
  menu.addItem('Session Management', 'showSessionManagement')
      .addSeparator();
  
  // System menu
  menu.addSubMenu(ui.createMenu('System')
      .addItem('Configuration', 'showConfiguration')
      .addItem('View Logs', 'showLogs')
      .addItem('Help', 'showHelp')
      .addItem('About', 'showAbout'));
  
  // Add menu to UI
  menu.addToUi();
}

/**
 * Creates a minimal emergency menu when regular initialization fails
 */
function createEmergencyMenu(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('YSLv6Hub Emergency')
    .addItem('Repair System', 'repairSystem')
    .addToUi();
}

/**
 * Initializes the system, checking for required sheets and creating them if needed
 */
function initializeSystem(): void {
  // Get the active spreadsheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Required sheet names
  const requiredSheets = [
    'YSLv6Hub',        // Dashboard
    'RegistrationInfo', // Registration data
    'SystemLog'         // System log
  ];
  
  // Verify/create required sheets
  for (const sheetName of requiredSheets) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      // Create the sheet if it doesn't exist
      sheet = ss.insertSheet(sheetName);
      
      // Initial setup based on sheet type
      setupNewSheet(sheet, sheetName);
    }
  }
  
  // Initialize system properties if needed
  initializeSystemProperties();
}

/**
 * Sets up a newly created sheet with appropriate formatting and headers
 */
function setupNewSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet, sheetName: string): void {
  switch (sheetName) {
    case 'YSLv6Hub':
      // Dashboard setup
      sheet.getRange('A1:C1').merge();
      sheet.getRange('A1').setValue('YSLv6Hub Dashboard')
        .setFontWeight('bold')
        .setFontSize(16);
      
      sheet.getRange('A3').setValue('Welcome to YSLv6Hub!');
      sheet.getRange('A4').setValue('This is the main dashboard for managing swim lessons.');
      sheet.getRange('A6').setValue('Getting Started:');
      sheet.getRange('A7').setValue('1. Import your registration data using the "Import Registration Data" menu');
      sheet.getRange('A8').setValue('2. Generate GroupsTracker sheets for your classes');
      sheet.getRange('A9').setValue('3. Use the sync functionality to keep your data in sync');
      
      // Adjust column widths
      sheet.setColumnWidth(1, 300);
      break;
      
    case 'RegistrationInfo':
      // Registration info setup
      const headers = [
        'First Name', 'Last Name', 'Age', 'Guardian', 'Email', 
        'Phone', 'Class', 'Stage', 'Schedule', 'Notes'
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
      
      // Freeze header row
      sheet.setFrozenRows(1);
      
      // Adjust column widths
      headers.forEach((_, index) => {
        sheet.setColumnWidth(index + 1, 120);
      });
      break;
      
    case 'SystemLog':
      // System log setup
      const logHeaders = [
        'Timestamp', 'Severity', 'Module', 'Function', 'Message', 'Details', 'User'
      ];
      sheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders])
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
      
      // Freeze header row
      sheet.setFrozenRows(1);
      
      // Adjust column widths
      sheet.setColumnWidth(1, 150); // Timestamp
      sheet.setColumnWidth(2, 80);  // Severity
      sheet.setColumnWidth(3, 120); // Module
      sheet.setColumnWidth(4, 120); // Function
      sheet.setColumnWidth(5, 300); // Message
      sheet.setColumnWidth(6, 200); // Details
      sheet.setColumnWidth(7, 150); // User
      break;
  }
}

/**
 * Initializes system properties if they don't already exist
 */
function initializeSystemProperties(): void {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Version
  if (!scriptProperties.getProperty('SYSTEM_VERSION')) {
    scriptProperties.setProperty('SYSTEM_VERSION', '6.0.0');
  }
  
  // Feature flags
  if (!scriptProperties.getProperty('FEATURE_FLAGS')) {
    const defaultFlags = {
      SMART_IMPORT: true,
      FORM_PROCESSOR: false,
      ENHANCED_COMMUNICATIONS: true,
      ADVANCED_REPORTING: false,
      SESSION_ANALYTICS: false
    };
    scriptProperties.setProperty('FEATURE_FLAGS', JSON.stringify(defaultFlags));
  }
}

/**
 * Shows the main dashboard
 */
function showDashboard(): void {
  // Get the dashboard sheet
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('YSLv6Hub');
  
  // If dashboard exists, activate it
  if (sheet) {
    sheet.activate();
  } else {
    // Otherwise create it
    const newSheet = ss.insertSheet('YSLv6Hub');
    setupNewSheet(newSheet, 'YSLv6Hub');
    newSheet.activate();
  }
}

/**
 * Repairs the system when issues occur
 */
function repairSystem(): void {
  try {
    // Reinitialize the system
    initializeSystem();
    
    // Recreate menu
    createMenu();
    
    // Show confirmation
    const ui = SpreadsheetApp.getUi();
    ui.alert('System Repair', 'YSLv6Hub system has been repaired successfully.', ui.ButtonSet.OK);
  } catch (error) {
    console.error('Failed to repair system:', String(error));
    const ui = SpreadsheetApp.getUi();
    ui.alert('System Repair Failed', 'Could not repair YSLv6Hub system. Please contact support.', ui.ButtonSet.OK);
  }
}

/**
 * Tests the initialization functionality
 * Used during development to verify functionality
 */
function testInitialization(): void {
  try {
    // Log start of test
    console.log('Testing YSLv6Hub initialization...');
    
    // Initialize system
    initializeSystem();
    
    // Check that required sheets exist
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requiredSheets = ['YSLv6Hub', 'RegistrationInfo', 'SystemLog'];
    const missingSheets = [];
    
    for (const sheetName of requiredSheets) {
      if (!ss.getSheetByName(sheetName)) {
        missingSheets.push(sheetName);
      }
    }
    
    // Check system properties
    const scriptProperties = PropertiesService.getScriptProperties();
    const systemVersion = scriptProperties.getProperty('SYSTEM_VERSION');
    const featureFlags = scriptProperties.getProperty('FEATURE_FLAGS');
    
    // Show results
    const ui = SpreadsheetApp.getUi();
    let message = 'Initialization Test Results:\n\n';
    
    if (missingSheets.length > 0) {
      message += `❌ Missing sheets: ${missingSheets.join(', ')}\n`;
    } else {
      message += '✅ All required sheets exist\n';
    }
    
    if (systemVersion) {
      message += `✅ System version: ${systemVersion}\n`;
    } else {
      message += '❌ System version not set\n';
    }
    
    if (featureFlags) {
      message += '✅ Feature flags configured\n';
    } else {
      message += '❌ Feature flags not configured\n';
    }
    
    message += '\nTest completed.';
    ui.alert('Initialization Test', message, ui.ButtonSet.OK);
    
    console.log('Initialization test completed');
  } catch (error) {
    console.error('Error in testInitialization:', error);
    const ui = SpreadsheetApp.getUi();
    ui.alert('Test Failed', `Initialization test failed: ${error instanceof Error ? error.message : String(error)}`, ui.ButtonSet.OK);
  }
}

// Placeholder functions for menu items
function importRegistrationData(): void {
  // Implementation will import registration data
  showNotImplemented('Import Registration Data');
}

function manageClasses(): void {
  // Implementation will manage classes
  showNotImplemented('Manage Classes');
}

function generateGroupsTracker(): void {
  // Implementation will generate groups tracker
  showNotImplemented('Generate GroupsTracker');
}

function syncData(): void {
  // Implementation will sync data
  showNotImplemented('Sync Data');
}

function showCommsHub(): void {
  // Implementation will show communications hub
  showNotImplemented('Communications Hub');
}

function generateReports(): void {
  // Implementation will generate reports
  showNotImplemented('Generate Reports');
}

function showGroupComms(): void {
  // Implementation will show group communications
  showNotImplemented('Group Communications');
}

function showInstructorComms(): void {
  // Implementation will show instructor communications
  showNotImplemented('Instructor Communications');
}

function showSessionManagement(): void {
  // Implementation will show session management
  showNotImplemented('Session Management');
}

function showConfiguration(): void {
  // Implementation will show configuration
  showNotImplemented('Configuration');
}

function showLogs(): void {
  // Implementation will show logs
  showNotImplemented('View Logs');
}

function showHelp(): void {
  // Implementation will show help
  showNotImplemented('Help');
}

function showAbout(): void {
  // Implementation will show about
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'About YSLv6Hub',
    'YSLv6Hub v6.0.0\n\nA comprehensive solution for managing YMCA swim lessons.\n\nDeveloped by: Sean Sullivan',
    ui.ButtonSet.OK
  );
}

/**
 * Shows a "not implemented" message for placeholder functions
 */
function showNotImplemented(feature: string): void {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Feature Not Implemented',
    `The "${feature}" feature is not yet implemented in this version.`,
    ui.ButtonSet.OK
  );
}

// Make functions available globally for menu actions
// @ts-ignore
global.onOpen = onOpen;
// @ts-ignore
global.showDashboard = showDashboard;
// @ts-ignore
global.repairSystem = repairSystem;
// @ts-ignore
global.testInitialization = testInitialization;
// @ts-ignore
global.importRegistrationData = importRegistrationData;
// @ts-ignore
global.manageClasses = manageClasses;
// @ts-ignore
global.generateGroupsTracker = generateGroupsTracker;
// @ts-ignore
global.syncData = syncData;
// @ts-ignore
global.showCommsHub = showCommsHub;
// @ts-ignore
global.generateReports = generateReports;
// @ts-ignore
global.showGroupComms = showGroupComms;
// @ts-ignore
global.showInstructorComms = showInstructorComms;
// @ts-ignore
global.showSessionManagement = showSessionManagement;
// @ts-ignore
global.showConfiguration = showConfiguration;
// @ts-ignore
global.showLogs = showLogs;
// @ts-ignore
global.showHelp = showHelp;
// @ts-ignore
global.showAbout = showAbout;