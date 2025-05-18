/**
 * YSL Hub v2 Version Control Actions Module
 * 
 * This module provides additional functionality for version control operations
 * beyond the basic version information tracking. It includes diagnostic tools
 * and update management functions that integrate with the VersionControl module.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

/**
 * Performs a comprehensive diagnostic check of the system
 * 
 * @returns Success status
 */
function performDiagnostics() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Performing system diagnostics', 'INFO', 'performDiagnostics');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Create a diagnostics sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let diagSheet = ss.getSheetByName('SystemDiagnostics');
    
    if (diagSheet) {
      // Clear existing content
      diagSheet.clear();
    } else {
      // Create new sheet
      diagSheet = ss.insertSheet('SystemDiagnostics');
    }
    
    // Set up sheet header
    diagSheet.getRange('A1:G1').merge()
      .setValue('YSL Hub System Diagnostics')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Get version information
    let versionInfo = 'Unknown';
    let releaseDate = 'Unknown';
    
    if (VersionControl && typeof VersionControl.getVersionInfo === 'function') {
      const vi = VersionControl.getVersionInfo();
      versionInfo = vi.currentVersion;
      releaseDate = vi.releaseDate;
    }
    
    diagSheet.getRange('A2:G2').merge()
      .setValue(`Version ${versionInfo} (${releaseDate})`)
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Add timestamp
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    diagSheet.getRange('A3:G3').merge()
      .setValue(`Diagnostic run: ${timestamp}`)
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Add section headers
    let currentRow = 5;
    
    // 1. System Configuration
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('1. System Configuration')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Get configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    // Check each configuration item
    const configChecks = [
      { name: 'Session Name', value: config.sessionName || 'Not set', required: true },
      { name: 'Roster Folder URL', value: config.rosterFolderUrl || 'Not set', required: true },
      { name: 'Report Template URL', value: config.reportTemplateUrl || 'Not set', required: true },
      { name: 'Swimmer Records URL', value: config.swimmerRecordsUrl || 'Not set', required: true },
      { name: 'Parent Handbook URL', value: config.parentHandbookUrl || 'Not set', required: false },
      { name: 'Session Programs URL', value: config.sessionProgramsUrl || 'Not set', required: true },
      { name: 'System Initialized', value: config.isInitialized ? 'Yes' : 'No', required: true }
    ];
    
    for (const check of configChecks) {
      diagSheet.getRange(currentRow, 1, 1, 2).merge().setValue(check.name);
      diagSheet.getRange(currentRow, 3, 1, 3).merge().setValue(check.value);
      
      // Add status indicator
      const status = check.required && (check.value === 'Not set' || check.value === 'No') ? 'ERROR' : 'OK';
      diagSheet.getRange(currentRow, 6, 1, 2).merge()
        .setValue(status)
        .setFontWeight('bold')
        .setBackground(status === 'OK' ? '#c6efce' : '#ffc7ce')
        .setFontColor(status === 'OK' ? '#006100' : '#9c0006');
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 2. Required Sheets
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('2. Required Sheets')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Check for required sheets
    const requiredSheets = [
      'Assumptions',
      'Classes',
      'Roster',
      'AssessmentCriteria',
      'Instructors',
      'SystemLog'
    ];
    
    for (const sheetName of requiredSheets) {
      const sheet = ss.getSheetByName(sheetName);
      diagSheet.getRange(currentRow, 1, 1, 2).merge().setValue(sheetName);
      
      // Add status indicator
      const status = sheet ? 'Present' : 'Missing';
      diagSheet.getRange(currentRow, 3, 1, 3).merge().setValue(status);
      diagSheet.getRange(currentRow, 6, 1, 2).merge()
        .setValue(sheet ? 'OK' : 'ERROR')
        .setFontWeight('bold')
        .setBackground(sheet ? '#c6efce' : '#ffc7ce')
        .setFontColor(sheet ? '#006100' : '#9c0006');
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 3. Script Properties
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('3. Script Properties')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Check critical script properties
    const scriptProps = PropertiesService.getScriptProperties();
    const properties = scriptProps.getProperties();
    
    // List of important properties to check
    const criticalProps = [
      'systemInitialized',
      'INITIALIZED',
      'sessionName',
      'SESSION_NAME',
      'swimmerRecordsUrl',
      'SWIMMER_RECORDS_URL'
    ];
    
    for (const prop of criticalProps) {
      const value = properties[prop] || 'Not set';
      diagSheet.getRange(currentRow, 1, 1, 2).merge().setValue(prop);
      diagSheet.getRange(currentRow, 3, 1, 3).merge().setValue(value);
      
      // Add status indicator
      const status = value === 'Not set' ? 'WARNING' : 'OK';
      diagSheet.getRange(currentRow, 6, 1, 2).merge()
        .setValue(status)
        .setFontWeight('bold')
        .setBackground(status === 'OK' ? '#c6efce' : '#fff2cc')
        .setFontColor(status === 'OK' ? '#006100' : '#9c6500');
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 4. Module Availability
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('4. Module Availability')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Check each module
    const modules = [
      { name: 'ErrorHandling', object: ErrorHandling },
      { name: 'GlobalFunctions', object: GlobalFunctions },
      { name: 'VersionControl', object: VersionControl },
      { name: 'AdministrativeModule', object: AdministrativeModule },
      { name: 'DataIntegrationModule', object: DataIntegrationModule },
      { name: 'CommunicationModule', object: CommunicationModule },
      { name: 'ReportingModule', object: ReportingModule },
      { name: 'UserGuide', object: UserGuide },
      { name: 'HistoryModule', object: HistoryModule },
      { name: 'SessionTransitionModule', object: SessionTransitionModule },
      { name: 'InstructorResourceModule', object: InstructorResourceModule },
      { name: 'DynamicInstructorSheet', object: DynamicInstructorSheet }
    ];
    
    for (const module of modules) {
      diagSheet.getRange(currentRow, 1, 1, 2).merge().setValue(module.name);
      
      // Check if module is defined
      const isAvailable = typeof module.object !== 'undefined';
      diagSheet.getRange(currentRow, 3, 1, 3).merge().setValue(isAvailable ? 'Available' : 'Missing');
      
      // Add status indicator
      diagSheet.getRange(currentRow, 6, 1, 2).merge()
        .setValue(isAvailable ? 'OK' : 'ERROR')
        .setFontWeight('bold')
        .setBackground(isAvailable ? '#c6efce' : '#ffc7ce')
        .setFontColor(isAvailable ? '#006100' : '#9c0006');
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 5. Menu Function Checks
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('5. Menu Function Checks')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Check key menu functions
    const menuFunctions = [
      { name: 'createFullMenu', func: createFullMenu },
      { name: 'YSL_SYNC_STUDENT_DATA', func: YSL_SYNC_STUDENT_DATA },
      { name: 'fixMenu', func: fixMenu },
      { name: 'installOnOpenTrigger', func: installOnOpenTrigger }
    ];
    
    for (const func of menuFunctions) {
      diagSheet.getRange(currentRow, 1, 1, 2).merge().setValue(func.name);
      
      // Check if function is defined
      const isAvailable = typeof func.func === 'function';
      diagSheet.getRange(currentRow, 3, 1, 3).merge().setValue(isAvailable ? 'Available' : 'Missing');
      
      // Add status indicator
      diagSheet.getRange(currentRow, 6, 1, 2).merge()
        .setValue(isAvailable ? 'OK' : 'ERROR')
        .setFontWeight('bold')
        .setBackground(isAvailable ? '#c6efce' : '#ffc7ce')
        .setFontColor(isAvailable ? '#006100' : '#9c0006');
      
      currentRow++;
    }
    
    // Add spacing
    currentRow++;
    
    // 6. Triggers
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('6. Installed Triggers')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Get all triggers
    const allTriggers = ScriptApp.getProjectTriggers();
    
    if (allTriggers.length === 0) {
      diagSheet.getRange(currentRow, 1, 1, 5).merge().setValue('No triggers installed');
      diagSheet.getRange(currentRow, 6, 1, 2).merge()
        .setValue('WARNING')
        .setFontWeight('bold')
        .setBackground('#fff2cc')
        .setFontColor('#9c6500');
      
      currentRow++;
    } else {
      // Add headers
      diagSheet.getRange(currentRow, 1).setValue('Function');
      diagSheet.getRange(currentRow, 2).setValue('Event Type');
      diagSheet.getRange(currentRow, 3).setValue('Source');
      diagSheet.getRange(currentRow, 4, 1, 4).merge().setValue('Status');
      
      diagSheet.getRange(currentRow, 1, 1, 7).setFontWeight('bold');
      
      currentRow++;
      
      // List triggers
      for (const trigger of allTriggers) {
        const handlerFunction = trigger.getHandlerFunction();
        const eventType = trigger.getEventType().toString();
        const triggerSource = trigger.getTriggerSource().toString();
        
        diagSheet.getRange(currentRow, 1).setValue(handlerFunction);
        diagSheet.getRange(currentRow, 2).setValue(eventType);
        diagSheet.getRange(currentRow, 3).setValue(triggerSource);
        
        // Check if critical onOpen trigger exists
        const isOnOpen = handlerFunction === 'onOpen' && 
                         eventType === ScriptApp.EventType.ON_OPEN.toString();
        
        diagSheet.getRange(currentRow, 4, 1, 4).merge()
          .setValue(isOnOpen ? 'OK - Required trigger' : 'OK')
          .setBackground('#c6efce')
          .setFontColor('#006100');
        
        currentRow++;
      }
      
      // Check if onOpen trigger is missing
      let hasOnOpenTrigger = false;
      for (const trigger of allTriggers) {
        if (trigger.getHandlerFunction() === 'onOpen' && 
            trigger.getEventType() === ScriptApp.EventType.ON_OPEN) {
          hasOnOpenTrigger = true;
          break;
        }
      }
      
      if (!hasOnOpenTrigger) {
        diagSheet.getRange(currentRow, 1).setValue('onOpen');
        diagSheet.getRange(currentRow, 2).setValue(ScriptApp.EventType.ON_OPEN.toString());
        diagSheet.getRange(currentRow, 3).setValue('SPREADSHEET');
        diagSheet.getRange(currentRow, 4, 1, 4).merge()
          .setValue('ERROR - Missing required trigger')
          .setBackground('#ffc7ce')
          .setFontColor('#9c0006');
        
        currentRow++;
      }
    }
    
    // Add spacing
    currentRow++;
    
    // 7. Summary and Recommendations
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('7. Summary and Recommendations')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    // Add overall status
    let overallStatus = 'System is functioning properly';
    let recommendations = '';
    
    // Check for specific issues
    if (!config.isInitialized) {
      overallStatus = 'System is not properly initialized';
      recommendations += '• Complete system initialization\n';
    }
    
    if (!config.swimmerRecordsUrl) {
      recommendations += '• Set Swimmer Records URL in configuration\n';
    }
    
    if (!config.rosterFolderUrl) {
      recommendations += '• Set Roster Folder URL in configuration\n';
    }
    
    let hasOnOpenTrigger = false;
    for (const trigger of allTriggers) {
      if (trigger.getHandlerFunction() === 'onOpen' && 
          trigger.getEventType() === ScriptApp.EventType.ON_OPEN) {
        hasOnOpenTrigger = true;
        break;
      }
    }
    
    if (!hasOnOpenTrigger) {
      recommendations += '• Install onOpen trigger using the installOnOpenTrigger function\n';
    }
    
    // Check for required sheets
    for (const sheetName of requiredSheets) {
      if (!ss.getSheetByName(sheetName)) {
        recommendations += `• Create missing sheet: ${sheetName}\n`;
      }
    }
    
    // Check for module objects
    for (const module of modules) {
      if (typeof module.object === 'undefined') {
        recommendations += `• Fix missing module: ${module.name}\n`;
      }
    }
    
    // Add summary to sheet
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue(overallStatus)
      .setFontWeight('bold');
    
    currentRow++;
    
    if (recommendations) {
      diagSheet.getRange(currentRow, 1, 1, 7).merge()
        .setValue('Recommendations:')
        .setFontWeight('bold');
      
      currentRow++;
      
      diagSheet.getRange(currentRow, 1, 1, 7).merge()
        .setValue(recommendations)
        .setWrap(true);
      
      // Set height to accommodate multiple lines
      diagSheet.setRowHeight(currentRow, 20 * (recommendations.split('\n').length + 1));
      
      currentRow += recommendations.split('\n').length;
    }
    
    // Add final repair options
    currentRow += 2;
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('Repair Options')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    currentRow++;
    
    diagSheet.getRange(currentRow, 1, 1, 7).merge()
      .setValue('If you are experiencing issues with the system, you can try the following repair options:')
      .setWrap(true);
    
    currentRow++;
    
    const repairOptions = [
      '1. Fix Menu: Use the fixMenu() function to reset menu properties and recreate menus',
      '2. Install Trigger: Use the installOnOpenTrigger() function to ensure the onOpen trigger is set up correctly',
      '3. Fix Initialization: Use AdministrativeModule_fixSystemInitializationProperty() to fix initialization status',
      '4. Clear Cache: Use VersionControl_clearCache() to clear any cached data that might be causing issues'
    ];
    
    for (const option of repairOptions) {
      diagSheet.getRange(currentRow, 1, 1, 7).merge().setValue(option);
      currentRow++;
    }
    
    // Format sheet
    diagSheet.setColumnWidth(1, 120);
    diagSheet.setColumnWidth(2, 120);
    diagSheet.setColumnWidth(3, 100);
    diagSheet.setColumnWidth(4, 100);
    diagSheet.setColumnWidth(5, 100);
    diagSheet.setColumnWidth(6, 80);
    diagSheet.setColumnWidth(7, 80);
    
    // Activate the diagnostics sheet
    diagSheet.activate();
    
    // Log the diagnostic run
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('System diagnostics completed', 'INFO', 'performDiagnostics');
    }
    
    // Show completion message
    ui.alert(
      'Diagnostics Complete',
      'System diagnostics have been completed. Review the SystemDiagnostics sheet for results and recommendations.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'performDiagnostics', 
        'Error performing system diagnostics. Please try again or contact support.');
    } else {
      Logger.log(`Error performing diagnostics: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to perform diagnostics: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Clears the system cache to resolve potential issues
 * 
 * @returns Success status
 */
function clearSystemCache() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Clearing system cache', 'INFO', 'clearSystemCache');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Confirm with user
    const result = ui.alert(
      'Clear System Cache',
      'This will clear all cached data and reset some properties. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return false;
    }
    
    // Clear cache properties
    const scriptProps = PropertiesService.getScriptProperties();
    const propertiesToClear = [
      'lastRosterSync',
      'lastAssessmentSync',
      'cachedClassData',
      'cachedRosterData',
      'cachedInstructorData'
    ];
    
    for (const prop of propertiesToClear) {
      scriptProps.deleteProperty(prop);
    }
    
    // Force menu refresh
    if (typeof reloadAllMenus === 'function') {
      reloadAllMenus();
    } else if (typeof fixMenu === 'function') {
      fixMenu();
    }
    
    // Show success message
    ui.alert(
      'Cache Cleared',
      'System cache has been cleared successfully. The menus have been refreshed.',
      ui.ButtonSet.OK
    );
    
    // Log the action
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('System cache cleared successfully', 'INFO', 'clearSystemCache');
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'clearSystemCache', 
        'Error clearing system cache. Please try again or contact support.');
    } else {
      Logger.log(`Error clearing system cache: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to clear system cache: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

// Global variable export - available to VersionControl module
const VersionControlActions = {
  performDiagnostics,
  clearSystemCache
};