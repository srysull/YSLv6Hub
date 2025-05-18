/**
 * YSL Hub v2 Debug Module
 * 
 * This module provides debugging tools and utilities for troubleshooting
 * and maintaining the YSL Hub system. It helps identify issues with the
 * system configuration, data integrity, and functionality.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

/**
 * Performs a system health check and displays status
 * 
 * @returns Success status
 */
function performSystemHealthCheck() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Performing system health check', 'INFO', 'performSystemHealthCheck');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Create a health check sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let healthSheet = ss.getSheetByName('SystemHealthCheck');
    
    if (healthSheet) {
      // Clear existing content
      healthSheet.clear();
    } else {
      // Create new sheet
      healthSheet = ss.insertSheet('SystemHealthCheck');
    }
    
    // Set up sheet header
    healthSheet.getRange('A1:E1').merge()
      .setValue('YSL Hub System Health Check')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Add timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    healthSheet.getRange('A2:E2').merge()
      .setValue(`Check performed: ${timestamp}`)
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Perform various health checks
    let currentRow = 4;
    
    // 1. Menu System Check
    healthSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('1. Menu System Check')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Add headers
    healthSheet.getRange(currentRow, 1).setValue('Component');
    healthSheet.getRange(currentRow, 2).setValue('Status');
    healthSheet.getRange(currentRow, 3, 1, 3).merge().setValue('Details');
    
    healthSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold');
    
    currentRow++;
    
    // Check key menu components
    const menuChecks = [
      { 
        name: 'createFullMenu function', 
        check: typeof createFullMenu === 'function',
        details: 'Creates the main YSL Hub menu structure'
      },
      { 
        name: 'AdministrativeModule.createMenu function', 
        check: AdministrativeModule && typeof AdministrativeModule.createMenu === 'function',
        details: 'Alternative menu creation function in AdministrativeModule'
      },
      {
        name: 'onOpen trigger',
        check: checkOnOpenTrigger(),
        details: 'Trigger to create menus when spreadsheet opens'
      },
      {
        name: 'System initialization status',
        check: PropertiesService.getScriptProperties().getProperty('systemInitialized') === 'true' ||
               PropertiesService.getScriptProperties().getProperty('INITIALIZED') === 'true',
        details: 'Required for full menu creation'
      }
    ];
    
    for (const check of menuChecks) {
      healthSheet.getRange(currentRow, 1).setValue(check.name);
      
      healthSheet.getRange(currentRow, 2)
        .setValue(check.check ? 'OK' : 'ISSUE')
        .setBackground(check.check ? '#c6efce' : '#ffc7ce')
        .setFontColor(check.check ? '#006100' : '#9c0006')
        .setFontWeight('bold');
      
      healthSheet.getRange(currentRow, 3, 1, 3).merge().setValue(check.details);
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 2. Data Structure Check
    healthSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('2. Data Structure Check')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Add headers
    healthSheet.getRange(currentRow, 1).setValue('Sheet');
    healthSheet.getRange(currentRow, 2).setValue('Status');
    healthSheet.getRange(currentRow, 3).setValue('Row Count');
    healthSheet.getRange(currentRow, 4, 1, 2).merge().setValue('Notes');
    
    healthSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold');
    
    currentRow++;
    
    // Check key data sheets
    const sheetChecks = [
      'Classes',
      'Roster',
      'AssessmentCriteria',
      'Instructors',
      'SystemLog',
      'Group Lesson Tracker'
    ];
    
    for (const sheetName of sheetChecks) {
      const sheet = ss.getSheetByName(sheetName);
      healthSheet.getRange(currentRow, 1).setValue(sheetName);
      
      if (sheet) {
        const rowCount = sheet.getLastRow();
        const isPopulated = rowCount > 1; // More than just a header row
        
        healthSheet.getRange(currentRow, 2)
          .setValue(isPopulated ? 'POPULATED' : 'EMPTY')
          .setBackground(isPopulated ? '#c6efce' : '#fff2cc')
          .setFontColor(isPopulated ? '#006100' : '#9c6500')
          .setFontWeight('bold');
        
        healthSheet.getRange(currentRow, 3).setValue(rowCount);
        
        let notes = '';
        if (sheetName === 'Group Lesson Tracker' && !isPopulated) {
          notes = 'Use "Generate Group Lesson Tracker" to create';
        } else if (!isPopulated) {
          notes = 'May need data refresh or import';
        }
        
        healthSheet.getRange(currentRow, 4, 1, 2).merge().setValue(notes);
      } else {
        healthSheet.getRange(currentRow, 2)
          .setValue('MISSING')
          .setBackground('#ffc7ce')
          .setFontColor('#9c0006')
          .setFontWeight('bold');
        
        healthSheet.getRange(currentRow, 3).setValue('N/A');
        
        let notes = '';
        if (sheetName === 'Group Lesson Tracker') {
          notes = 'Use "Generate Group Lesson Tracker" to create';
        } else {
          notes = 'Required sheet is missing';
        }
        
        healthSheet.getRange(currentRow, 4, 1, 2).merge().setValue(notes);
      }
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 3. Module Check
    healthSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('3. Module Check')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Add headers
    healthSheet.getRange(currentRow, 1).setValue('Module');
    healthSheet.getRange(currentRow, 2).setValue('Status');
    healthSheet.getRange(currentRow, 3, 1, 3).merge().setValue('Impact');
    
    healthSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold');
    
    currentRow++;
    
    // Check each module
    const moduleChecks = [
      { 
        name: 'ErrorHandling', 
        check: typeof ErrorHandling !== 'undefined',
        impact: 'Logging, error management, and user feedback'
      },
      { 
        name: 'GlobalFunctions', 
        check: typeof GlobalFunctions !== 'undefined',
        impact: 'Core functionality and utilities'
      },
      { 
        name: 'VersionControl', 
        check: typeof VersionControl !== 'undefined',
        impact: 'Version tracking and system information'
      },
      { 
        name: 'AdministrativeModule', 
        check: typeof AdministrativeModule !== 'undefined',
        impact: 'System initialization and configuration'
      },
      { 
        name: 'DataIntegrationModule', 
        check: typeof DataIntegrationModule !== 'undefined',
        impact: 'Data synchronization and management'
      },
      { 
        name: 'CommunicationModule', 
        check: typeof CommunicationModule !== 'undefined',
        impact: 'Email communications with parents'
      },
      { 
        name: 'ReportingModule', 
        check: typeof ReportingModule !== 'undefined',
        impact: 'Progress and assessment reports'
      },
      { 
        name: 'UserGuide', 
        check: typeof UserGuide !== 'undefined',
        impact: 'System documentation'
      },
      { 
        name: 'HistoryModule', 
        check: typeof HistoryModule !== 'undefined',
        impact: 'System event tracking'
      },
      { 
        name: 'SessionTransitionModule', 
        check: typeof SessionTransitionModule !== 'undefined',
        impact: 'Session transitions'
      },
      { 
        name: 'InstructorResourceModule', 
        check: typeof InstructorResourceModule !== 'undefined',
        impact: 'Instructor sheets and resources'
      },
      { 
        name: 'DynamicInstructorSheet', 
        check: typeof DynamicInstructorSheet !== 'undefined',
        impact: 'Group Lesson Tracker creation'
      },
      { 
        name: 'VersionControlActions', 
        check: typeof VersionControlActions !== 'undefined',
        impact: 'System diagnostics and cache management'
      },
      { 
        name: 'DebugModule', 
        check: typeof DebugModule !== 'undefined',
        impact: 'Debugging and troubleshooting tools'
      }
    ];
    
    for (const module of moduleChecks) {
      healthSheet.getRange(currentRow, 1).setValue(module.name);
      
      healthSheet.getRange(currentRow, 2)
        .setValue(module.check ? 'LOADED' : 'MISSING')
        .setBackground(module.check ? '#c6efce' : '#ffc7ce')
        .setFontColor(module.check ? '#006100' : '#9c0006')
        .setFontWeight('bold');
      
      healthSheet.getRange(currentRow, 3, 1, 3).merge().setValue(module.impact);
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 4. Configuration Check
    healthSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('4. Configuration Check')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Get configuration
    const config = AdministrativeModule && typeof AdministrativeModule.getSystemConfiguration === 'function' ?
      AdministrativeModule.getSystemConfiguration() : {};
    
    // Add headers
    healthSheet.getRange(currentRow, 1).setValue('Setting');
    healthSheet.getRange(currentRow, 2).setValue('Status');
    healthSheet.getRange(currentRow, 3, 1, 3).merge().setValue('Value/Notes');
    
    healthSheet.getRange(currentRow, 1, 1, 5).setFontWeight('bold');
    
    currentRow++;
    
    // Configuration checks
    const configChecks = [
      { 
        name: 'System Initialized', 
        check: config.isInitialized === true,
        value: config.isInitialized ? 'Yes' : 'No'
      },
      { 
        name: 'Session Name', 
        check: Boolean(config.sessionName),
        value: config.sessionName || 'Not set'
      },
      { 
        name: 'Roster Folder URL', 
        check: Boolean(config.rosterFolderUrl),
        value: config.rosterFolderUrl || 'Not set'
      },
      { 
        name: 'Swimmer Records URL', 
        check: Boolean(config.swimmerRecordsUrl),
        value: config.swimmerRecordsUrl || 'Not set'
      },
      { 
        name: 'Report Template URL', 
        check: Boolean(config.reportTemplateUrl),
        value: config.reportTemplateUrl || 'Not set'
      },
      { 
        name: 'Session Programs URL', 
        check: Boolean(config.sessionProgramsUrl),
        value: config.sessionProgramsUrl || 'Not set'
      }
    ];
    
    for (const check of configChecks) {
      healthSheet.getRange(currentRow, 1).setValue(check.name);
      
      healthSheet.getRange(currentRow, 2)
        .setValue(check.check ? 'OK' : 'MISSING')
        .setBackground(check.check ? '#c6efce' : '#ffc7ce')
        .setFontColor(check.check ? '#006100' : '#9c0006')
        .setFontWeight('bold');
      
      healthSheet.getRange(currentRow, 3, 1, 3).merge().setValue(check.value);
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 5. Repair Actions
    healthSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('5. Repair Actions')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    healthSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('If you are experiencing issues with the system, try the following repair actions:')
      .setWrap(true);
    
    currentRow++;
    
    const repairActions = [
      'Use fixMenu() to recreate the menu system',
      'Use installOnOpenTrigger() to ensure the onOpen trigger is properly installed',
      'Use AdministrativeModule_fixSystemInitializationProperty() to fix initialization issues',
      'Use VersionControlActions.clearSystemCache() to clear cached data'
    ];
    
    for (const action of repairActions) {
      healthSheet.getRange(currentRow, 1, 1, 5).merge().setValue(`• ${action}`);
      currentRow++;
    }
    
    // Format the sheet
    healthSheet.setColumnWidth(1, 200);
    healthSheet.setColumnWidth(2, 100);
    healthSheet.setColumnWidth(3, 100);
    healthSheet.setColumnWidth(4, 150);
    healthSheet.setColumnWidth(5, 150);
    
    // Activate the health check sheet
    healthSheet.activate();
    
    // Log the health check completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('System health check completed', 'INFO', 'performSystemHealthCheck');
    }
    
    // Show completion message
    ui.alert(
      'Health Check Complete',
      'System health check has been completed. Review the SystemHealthCheck sheet for results and recommendations.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'performSystemHealthCheck', 
        'Error performing system health check. Please try again or contact support.');
    } else {
      Logger.log(`Error performing health check: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to perform health check: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Repairs the system by resetting various components
 * 
 * @returns Success status
 */
function repairSystem() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Repairing system', 'INFO', 'repairSystem');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Confirm with user
    const result = ui.alert(
      'Repair System',
      'This will attempt to repair common system issues by resetting various components. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return false;
    }
    
    // 1. Fix system initialization properties
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('systemInitialized', 'true');
    scriptProps.setProperty('INITIALIZED', 'true');
    
    // 2. Install onOpen trigger
    if (typeof installOnOpenTrigger === 'function') {
      installOnOpenTrigger();
    }
    
    // 3. Fix menu
    if (typeof fixMenu === 'function') {
      fixMenu();
    } else if (typeof reloadAllMenus === 'function') {
      reloadAllMenus();
    }
    
    // 4. Clear system cache
    if (VersionControlActions && typeof VersionControlActions.clearSystemCache === 'function') {
      VersionControlActions.clearSystemCache();
    } else {
      // Manual cache clearing
      const cacheProps = [
        'lastRosterSync',
        'lastAssessmentSync',
        'cachedClassData',
        'cachedRosterData',
        'cachedInstructorData'
      ];
      
      for (const prop of cacheProps) {
        scriptProps.deleteProperty(prop);
      }
    }
    
    // 5. Check for and repair data structures
    if (DataIntegrationModule && typeof DataIntegrationModule.initializeDataStructures === 'function') {
      // Get current config
      const config = AdministrativeModule && typeof AdministrativeModule.getSystemConfiguration === 'function' ?
        AdministrativeModule.getSystemConfiguration() : {};
      
      DataIntegrationModule.initializeDataStructures(config);
    }
    
    // Show success message
    ui.alert(
      'Repair Complete',
      'System repair has been completed. Please refresh the spreadsheet to see if the issues have been resolved.',
      ui.ButtonSet.OK
    );
    
    // Log the repair completion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('System repair completed', 'INFO', 'repairSystem');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'repairSystem', 
        'Error repairing system. Please try again or contact support.');
    } else {
      Logger.log(`Error repairing system: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to repair system: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Checks if onOpen trigger is installed
 * 
 * @returns True if trigger exists, false otherwise
 */
function checkOnOpenTrigger() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === 'onOpen' && 
          trigger.getEventType() === ScriptApp.EventType.ON_OPEN) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    Logger.log(`Error checking onOpen trigger: ${error.message}`);
    return false;
  }
}

/**
 * Tests menu creation by creating a diagnostic menu
 * 
 * @returns Success status
 */
function testMenuCreation() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Testing menu creation', 'INFO', 'testMenuCreation');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Create a simple test menu
    ui.createMenu('Debug Menu')
      .addItem('System Health Check', 'DebugModule.performSystemHealthCheck')
      .addItem('Repair System', 'DebugModule.repairSystem')
      .addItem('Fix Menu', 'fixMenu')
      .addItem('Install onOpen Trigger', 'installOnOpenTrigger')
      .addToUi();
    
    // Show success message
    ui.alert(
      'Test Menu Created',
      'A test Debug menu has been created successfully. If you can see this menu, the menu creation system is working.',
      ui.ButtonSet.OK
    );
    
    // Log the test
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Menu creation test successful', 'INFO', 'testMenuCreation');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'testMenuCreation', 
        'Error testing menu creation. Menu system may be broken.');
    } else {
      Logger.log(`Error testing menu creation: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to test menu creation: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

// Global variable export
const DebugModule = {
  performSystemHealthCheck,
  repairSystem,
  checkOnOpenTrigger,
  testMenuCreation
};