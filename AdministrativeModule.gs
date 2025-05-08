  /**
 * YSL Hub v2 Administrative Module
 * 
 * This module handles system initialization, configuration management, and user interface
 * elements for the YSL Hub system. It serves as the primary entry point for user interactions
 * and coordinates the initialization workflow.
 * 
 * @author PenBay YMCA
 * @version 2.0
 * @date 2025-04-27
 */

// Configuration property keys
const CONFIG = {
  SESSION_NAME: 'sessionName',
  ROSTER_FOLDER_URL: 'rosterFolderUrl',
  REPORT_TEMPLATE_URL: 'reportTemplateUrl',
  SWIMMER_RECORDS_URL: 'swimmerRecordsUrl',
  PARENT_HANDBOOK_URL: 'parentHandbookUrl',
  SESSION_PROGRAMS_URL: 'sessionProgramsUrl',
  INITIALIZED: 'systemInitialized'
};

// Default folder IDs for critical system resources
const DEFAULT_FOLDERS = {
  SESSION_ROSTERS: '1vlR8WwEyLWOuO-JUzCzrdLikv-hVVld4',
  REPORT_TEMPLATES: '',
  DIVAB_INTAKE: '1_Zag1Qlb_jbwZ15URuZzzdSBBNgYwih_yBcnRS_VbeY'
};

/**
 * Creates the YSL Hub menu in the spreadsheet UI.
 * Menu structure changes based on system initialization status.
 */
function createMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('YSL Hub');
    
    // Check if system is initialized
    const isInitialized = GlobalFunctions.safeGetProperty('systemInitialized') === 'true';
    
    // Check if initialization has been started but not completed
    const sessionName = GlobalFunctions.safeGetProperty('sessionName');
    const hasStartedInit = sessionName && !isInitialized;
    
    // Default menu items that will always be added
    let hasItems = false;
    
    if (!isInitialized) {
      if (hasStartedInit) {
        // System initialization started but not completed - show Full Initialization option
        menu.addItem('Initialize System', 'AdministrativeModule.showInitializationDialog')
            .addItem('Full Initialization', 'AdministrativeModule.fullInitialization')
            .addSeparator()
            .addItem('About YSL Hub', 'AdministrativeModule.showAboutDialog');
        hasItems = true;
      } else {
        // System not initialized at all - just show initialization option
        menu.addItem('Initialize System', 'AdministrativeModule.showInitializationDialog')
            .addSeparator()
            .addItem('About YSL Hub', 'AdministrativeModule.showAboutDialog');
        hasItems = true;
      }
    } else {
      // System is initialized - show operational options
      try {
        menu.addSubMenu(ui.createMenu('Class Management')
              .addItem('Update Class Selector', 'DataIntegrationModule_updateClassSelector')
              .addItem('Refresh Roster Data', 'DataIntegrationModule_refreshRosterData')
              .addItem('Update Instructor Data', 'DataIntegrationModule_updateInstructorData')
              .addItem('Generate Instructor Sheets', 'InstructorResourceModule_generateInstructorSheets')
              .addSeparator()
              .addItem('Create Dynamic Instructor Sheet', 'DynamicInstructorSheet.createDynamicInstructorSheet')
              .addItem('Push Skills to Swimmer Records', 'DynamicInstructorSheet.pushSkillsToSwimmerRecords'))
            .addSubMenu(ui.createMenu('Assessment Management')
              .addItem('Push Assessments to Swimmer Log', 'DataIntegrationModule_pushAssessmentsToSwimmerLog')
              .addItem('Pull Latest Assessment Criteria', 'DataIntegrationModule_pullAssessmentCriteria'))
            .addSubMenu(ui.createMenu('Reports')
              .addItem('Report Assessment Criteria', 'DataIntegrationModule_reportAssessmentCriteria')
              .addItem('Diagnose Criteria Import', 'DataIntegrationModule_diagnoseCriteriaImport')
              .addItem('Generate Mid-Session Reports', 'ReportingModule_generateMidSessionReports')
              .addItem('Generate End-Session Reports', 'ReportingModule_generateEndSessionReports'))
            .addSubMenu(ui.createMenu('Communications')
              .addItem('Email Class Participants', 'CommunicationModule_emailClassParticipants')
              .addItem('Send Class Announcements', 'CommunicationModule_sendClassAnnouncements')
              .addItem('Test Welcome Email', 'CommunicationModule.testWelcomeEmail')
              .addItem('Send Welcome Email', 'CommunicationModule.sendWelcomeEmails'))
            .addSubMenu(ui.createMenu('System')
              .addItem('System Configuration', 'AdministrativeModule_showConfigurationDialog')
              .addItem('View System Log', 'ErrorHandling_showLogViewer')
              .addItem('Hide System Log', 'ErrorHandling_hideLogSheet')
              .addItem('Clear System Log', 'ErrorHandling_clearLog')
              .addItem('Export System Log', 'ErrorHandling_exportLog')
              .addItem('Run System Diagnostics', 'VersionControl_showDiagnostics')
              .addItem('View Version Information', 'VersionControl_showVersionInfo'))
            .addSeparator()
            .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog');
        hasItems = true;
      } catch (menuError) {
        // If menu creation fails, log it but continue with fallback menu
        Logger.log(`Error creating operational menu: ${menuError.message}`);
        // The operational menu creation failed, so we'll add fallback items
        hasItems = false;
      }
    }
    
    // If no menu items were added (which would cause an error), add fallback items
    if (!hasItems) {
      menu.addItem('System Configuration', 'AdministrativeModule_showConfigurationDialog')
          .addItem('About YSL Hub', 'AdministrativeModule_showAboutDialog');
    }
    
    // Add the menu to the UI
    menu.addToUi();
    return menu;
  } catch (error) {
    // Use basic error logging since ErrorHandling might not be initialized yet
    Logger.log(`Error creating menu: ${error.message}`);
    
    // Try to create a minimal menu as a last resort
    try {
      const ui = SpreadsheetApp.getUi();
      ui.createMenu('YSL Hub')
        .addItem('Repair System', 'AdministrativeModule_showConfigurationDialog')
        .addToUi();
    } catch (finalError) {
      Logger.log(`Failed to create minimal menu: ${finalError.message}`);
    }
    
    return null;
  }
}

/**
 * Displays the initialization dialog to gather initial configuration parameters.
 */
/**
 * Shows the initialization dialog to gather all required configuration parameters.
 * @return {boolean} Success status
 */
function showInitializationDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting guided initialization dialog', 'INFO', 'showInitializationDialog');
    }
    
    // Step 1: Get session name
    const sessionResult = ui.prompt(
      'YSL Hub Initialization (Step 1 of 4)',
      'Please enter the session name (e.g., "Spring2 2025"):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (sessionResult.getSelectedButton() !== ui.Button.OK) return false;
    const sessionName = sessionResult.getResponseText().trim();
    
    if (!sessionName) {
      ui.alert('Error', 'Session name cannot be empty. Please try again.', ui.ButtonSet.OK);
      return false;
    }
    
    // Step 2: Get report template folder URL
    const reportTemplateResult = ui.prompt(
      'YSL Hub Initialization (Step 2 of 4)',
      'Please enter the URL of the report template folder in Google Drive:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (reportTemplateResult.getSelectedButton() !== ui.Button.OK) return false;
    const reportTemplateUrl = reportTemplateResult.getResponseText().trim();
    
    if (!reportTemplateUrl) {
      ui.alert('Error', 'Report template folder URL cannot be empty. Please try again.', ui.ButtonSet.OK);
      return false;
    }
    
    // Step 3: Get swimmer records workbook URL
    const swimmerRecordsResult = ui.prompt(
      'YSL Hub Initialization (Step 3 of 4)',
      'Please enter the URL of the YSL Swimmer Records workbook:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (swimmerRecordsResult.getSelectedButton() !== ui.Button.OK) return false;
    const swimmerRecordsUrl = swimmerRecordsResult.getResponseText().trim();
    
    if (!swimmerRecordsUrl) {
      ui.alert('Error', 'Swimmer records workbook URL cannot be empty. Please try again.', ui.ButtonSet.OK);
      return false;
    }
    
    // Step 4: Get session programs workbook URL
    const sessionProgramsResult = ui.prompt(
      'YSL Hub Initialization (Step 4 of 4)',
      'Please enter the URL of the Session Programs workbook:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (sessionProgramsResult.getSelectedButton() !== ui.Button.OK) return false;
    const sessionProgramsUrl = sessionProgramsResult.getResponseText().trim();
    
    if (!sessionProgramsUrl) {
      ui.alert('Error', 'Session programs workbook URL cannot be empty. Please try again.', ui.ButtonSet.OK);
      return false;
    }
    
    // Optional: Get parent handbook URL
    const parentHandbookResult = ui.prompt(
      'YSL Hub Initialization (Optional)',
      'Please enter the URL of the Parent Handbook PDF (optional, can be left blank):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (parentHandbookResult.getSelectedButton() !== ui.Button.OK) return false;
    const parentHandbookUrl = parentHandbookResult.getResponseText().trim();
    
    // Store all configuration values in script properties
    GlobalFunctions.safeSetProperty(CONFIG.SESSION_NAME, sessionName);
    GlobalFunctions.safeSetProperty(CONFIG.ROSTER_FOLDER_URL, `https://drive.google.com/drive/folders/${DEFAULT_FOLDERS.SESSION_ROSTERS}`);
    GlobalFunctions.safeSetProperty(CONFIG.REPORT_TEMPLATE_URL, reportTemplateUrl);
    GlobalFunctions.safeSetProperty(CONFIG.SWIMMER_RECORDS_URL, swimmerRecordsUrl);
    GlobalFunctions.safeSetProperty(CONFIG.SESSION_PROGRAMS_URL, sessionProgramsUrl);
    
    if (parentHandbookUrl) {
      GlobalFunctions.safeSetProperty(CONFIG.PARENT_HANDBOOK_URL, parentHandbookUrl);
    }
    
    // Prepare the Assumptions sheet with the provided values
    prepareAssumptionsSheetWithValues(sessionName, reportTemplateUrl, swimmerRecordsUrl, sessionProgramsUrl, parentHandbookUrl);
    
    // Combine all values into a configuration object for initialization
    const configValues = {
      sessionName: sessionName,
      rosterFolderUrl: `https://drive.google.com/drive/folders/${DEFAULT_FOLDERS.SESSION_ROSTERS}`,
      reportTemplateUrl: reportTemplateUrl,
      swimmerRecordsUrl: swimmerRecordsUrl,
      parentHandbookUrl: parentHandbookUrl,
      sessionProgramsUrl: sessionProgramsUrl
    };
    
    // Log the values
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Configuration values collected: sessionName=${configValues.sessionName}, rosterFolderUrl=${configValues.rosterFolderUrl}, reportTemplateUrl=${configValues.reportTemplateUrl}, swimmerRecordsUrl=${configValues.swimmerRecordsUrl}, sessionProgramsUrl=${configValues.sessionProgramsUrl}`, 'INFO', 'showInitializationDialog');
    }
    
    // Confirm and perform initialization
    const confirmResult = ui.alert(
      'Ready to Initialize',
      'All required configuration has been collected. Would you like to proceed with system initialization?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResult !== ui.Button.YES) {
      ui.alert(
        'Initialization Paused',
        'System configuration has been saved. You can complete initialization later by selecting "Complete Initialization" from the YSL Hub menu.',
        ui.ButtonSet.OK
      );
      
      // Update menu to show Complete Initialization option
      createMenu();
      return false;
    }
    
    // Perform the initialization
    return performInitialization(configValues);
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'showInitializationDialog', 
        'Error during initialization. Please try again or contact support.');
    } else {
      Logger.log(`Initialization dialog error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error during initialization: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Helper function to prepare Assumptions sheet with all collected values
 * @param {string} sessionName - The session name
 * @param {string} reportTemplateUrl - URL of the report template folder
 * @param {string} swimmerRecordsUrl - URL of the swimmer records workbook
 * @param {string} sessionProgramsUrl - URL of the session programs workbook
 * @param {string} parentHandbookUrl - URL of the parent handbook PDF (optional)
 */
function prepareAssumptionsSheetWithValues(sessionName, reportTemplateUrl, swimmerRecordsUrl, sessionProgramsUrl, parentHandbookUrl) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Assumptions');
    
    // If Assumptions sheet doesn't exist, rename the first sheet
    if (!sheet) {
      sheet = ss.getSheets()[0];
      sheet.setName('Assumptions');
    } else {
      // Clear existing content if sheet already exists
      sheet.clear();
    }
    
    // Set up headers and instructions
    sheet.getRange('A1:B1').merge()
      .setValue('YSL Hub Configuration')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Instructions
    sheet.getRange('A2:B5').merge()
      .setValue('YSL Hub has been configured with the values below. These settings can be modified through the System Configuration option in the YSL Hub menu.')
      .setWrap(true);
    
    // Configuration fields
    const configFields = [
      ['Session Name', sessionName],
      ['Session Roster Folder URL', `https://drive.google.com/drive/folders/${DEFAULT_FOLDERS.SESSION_ROSTERS}`],
      ['Report Template Folder URL', reportTemplateUrl],
      ['Swimmer Records Workbook URL', swimmerRecordsUrl],
      ['Parent Handbook PDF URL', parentHandbookUrl || ''],
      ['Session Programs Workbook URL', sessionProgramsUrl]
    ];
    
    // Add field labels and input cells
    const startRow = 7;
    configFields.forEach((field, index) => {
      const row = startRow + index;
      sheet.getRange(`A${row}`).setValue(field[0]).setFontWeight('bold');
      sheet.getRange(`B${row}`).setValue(field[1]);
    });
    
    // Add note about roster file naming convention
    const noteRow = startRow + configFields.length + 1;
    sheet.getRange(`A${noteRow}:B${noteRow}`).merge()
      .setValue('Note: The session roster file should follow the naming convention: ' +
                `"YSL ${sessionName} Roster"`)
      .setFontStyle('italic')
      .setWrap(true);
    
    // Format sheet
    sheet.autoResizeColumn(1);
    sheet.setColumnWidth(2, 400);
    
    // Log successful preparation
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Assumptions sheet prepared with all configuration values', 'INFO', 'prepareAssumptionsSheetWithValues');
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error preparing Assumptions sheet: ${error.message}`, 'ERROR', 'prepareAssumptionsSheetWithValues');
    } else {
      Logger.log(`Error preparing Assumptions sheet: ${error.message}`);
      throw error; // Re-throw to caller
    }
  }
}

/**
 * Performs the actual system initialization with the provided configuration values
 * @param {Object} configValues - The configuration values
 * @return {boolean} Success status
 */
function performInitialization(configValues) {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Initialize error handling system if not already initialized
    if (ErrorHandling && typeof ErrorHandling.initializeErrorHandling === 'function') {
      ErrorHandling.initializeErrorHandling();
    }
    
    // Log the initialization start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting system initialization', 'INFO', 'performInitialization');
    }
    
    // Initialize the system components
    DataIntegrationModule.initializeDataStructures(configValues);
    
    // Mark system as initialized
    GlobalFunctions.safeSetProperty(CONFIG.INITIALIZED, 'true');
    
    // Record initialization in version control
    if (VersionControl && typeof VersionControl.recordUpdate === 'function') {
      VersionControl.recordUpdate(VersionControl.getCurrentVersion(), 
        `System initialized for session: ${configValues.sessionName}`);
    }
    
    // Show success message
    ui.alert(
      'Initialization Complete',
      'YSL Hub has been successfully initialized. You can now use the system features from the YSL Hub menu.',
      ui.ButtonSet.OK
    );
    
    // Refresh menu to show operational options
    createMenu();
    
    // Log successful initialization
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('System initialization completed successfully', 'INFO', 'performInitialization');
    }
    
    return true;
  } catch (error) {
    // Handle initialization errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'performInitialization', 
        'Error during system initialization. Please check the configuration and try again.');
    } else {
      Logger.log(`Initialization error: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Initialization Failed',
        `An error occurred during initialization: ${error.message}. Please check the configuration and try again.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Prepares the Assumptions sheet with configuration parameters.
 * @param {string} sessionName - The name of the session
 */
function prepareAssumptionsSheet(sessionName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Assumptions');
    
    // If Assumptions sheet doesn't exist, rename the first sheet
    if (!sheet) {
      sheet = ss.getSheets()[0];
      sheet.setName('Assumptions');
    } else {
      // Clear existing content if sheet already exists
      sheet.clear();
    }
    
    // Set up headers and instructions
    sheet.getRange('A1:B1').merge()
      .setValue('YSL Hub Configuration')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Instructions
    sheet.getRange('A2:B5').merge()
      .setValue('Please provide the following information to complete system initialization. ' +
                'After filling in all fields, select "Full Initialization" from the YSL Hub menu.')
      .setWrap(true);
    
    // Configuration fields
    const configFields = [
      ['Session Name', sessionName],
      ['Session Roster Folder URL', `https://drive.google.com/drive/folders/${DEFAULT_FOLDERS.SESSION_ROSTERS}`],
      ['Report Template Folder URL', ''],
      ['Swimmer Records Workbook URL', ''],
      ['Parent Handbook PDF URL', ''],
      ['Session Programs Workbook URL', ''] // For instructor assignments
    ];
    
    // Add field labels and input cells
    const startRow = 7;
    configFields.forEach((field, index) => {
      const row = startRow + index;
      sheet.getRange(`A${row}`).setValue(field[0]).setFontWeight('bold');
      sheet.getRange(`B${row}`).setValue(field[1]);
    });
    
    // Add note about roster file naming convention
    const noteRow = startRow + configFields.length + 1;
    sheet.getRange(`A${noteRow}:B${noteRow}`).merge()
      .setValue('Note: The session roster file should follow the naming convention: ' +
                `"YSL ${sessionName} Roster"`)
      .setFontStyle('italic')
      .setWrap(true);
    
    // Add note about Session Programs workbook
    const programsNoteRow = noteRow + 1;
    sheet.getRange(`A${programsNoteRow}:B${programsNoteRow}`).merge()
      .setValue('Note: The Session Programs workbook should contain instructor assignments ' +
                'and will be used to populate instructor information in the Classes sheet.')
      .setFontStyle('italic')
      .setWrap(true);
    
    // Format sheet
    sheet.autoResizeColumn(1);
    sheet.setColumnWidth(2, 400);
    
    // Log successful preparation
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Assumptions sheet prepared successfully', 'INFO', 'prepareAssumptionsSheet');
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'prepareAssumptionsSheet', 
        'Error preparing Assumptions sheet. Please try again or contact support.');
    } else {
      Logger.log(`Error preparing Assumptions sheet: ${error.message}`);
      throw error; // Re-throw to caller
    }
  }
}

function fullInitialization() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assumptionsSheet = ss.getSheetByName('Assumptions');
    
    if (!assumptionsSheet) {
      ui.alert('Error', 'Assumptions sheet not found. Please restart initialization.', ui.ButtonSet.OK);
      return;
    }
    
    // Extract configuration parameters directly from specific cells 
    // to avoid potential issues with the getRange().getValues() method
    const configValues = {
      sessionName: assumptionsSheet.getRange('B7').getValue(),
      rosterFolderUrl: assumptionsSheet.getRange('B8').getValue(),
      reportTemplateUrl: assumptionsSheet.getRange('B9').getValue(),
      swimmerRecordsUrl: assumptionsSheet.getRange('B10').getValue(),
      parentHandbookUrl: assumptionsSheet.getRange('B11').getValue(),
      sessionProgramsUrl: assumptionsSheet.getRange('B12').getValue()
    };
    
    // Log the values we're finding to help debug
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Config values found using direct cell access: sessionName=${configValues.sessionName}, rosterFolderUrl=${configValues.rosterFolderUrl}, reportTemplateUrl=${configValues.reportTemplateUrl}, swimmerRecordsUrl=${configValues.swimmerRecordsUrl}, sessionProgramsUrl=${configValues.sessionProgramsUrl}`, 'INFO', 'fullInitialization');
    }
    
    // Ensure all values are properly trimmed
    const validatedValues = {
      sessionName: configValues.sessionName ? configValues.sessionName.toString().trim() : '',
      rosterFolderUrl: configValues.rosterFolderUrl ? configValues.rosterFolderUrl.toString().trim() : '',
      reportTemplateUrl: configValues.reportTemplateUrl ? configValues.reportTemplateUrl.toString().trim() : '',
      swimmerRecordsUrl: configValues.swimmerRecordsUrl ? configValues.swimmerRecordsUrl.toString().trim() : '',
      parentHandbookUrl: configValues.parentHandbookUrl ? configValues.parentHandbookUrl.toString().trim() : '',
      sessionProgramsUrl: configValues.sessionProgramsUrl ? configValues.sessionProgramsUrl.toString().trim() : ''
    };
    
    // Validation approach
    const requiredFields = ['sessionName', 'rosterFolderUrl', 'reportTemplateUrl', 'swimmerRecordsUrl', 'sessionProgramsUrl'];
    const missingFields = [];

    for (const field of requiredFields) {
      if (!validatedValues[field]) {
        missingFields.push(field);
      }
    }

    if (missingFields.length > 0) {
      ui.alert(
        'Missing Configuration',
        `Please fill in all required fields in the Assumptions sheet: ${missingFields.join(', ')}`,
        ui.ButtonSet.OK
      );
      return;
    }
    
    // Store configuration in script properties
    Object.entries(validatedValues).forEach(([key, value]) => {
      if (value) { // Only store non-empty values
        GlobalFunctions.safeSetProperty(CONFIG[key.toUpperCase()], value);
      }
    });
    
    // Rest of the function remains the same...
    
    // Initialize error handling system if not already initialized
    if (ErrorHandling && typeof ErrorHandling.initializeErrorHandling === 'function') {
      ErrorHandling.initializeErrorHandling();
    }
    
    // Log the initialization start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting full system initialization', 'INFO', 'fullInitialization');
    }
    
    // Initialize the system components
    DataIntegrationModule.initializeDataStructures(validatedValues);
    
    // Mark system as initialized
    GlobalFunctions.safeSetProperty(CONFIG.INITIALIZED, 'true');
    
    // Record initialization in version control
    if (VersionControl && typeof VersionControl.recordUpdate === 'function') {
      VersionControl.recordUpdate(VersionControl.getCurrentVersion(), 
        `System initialized for session: ${validatedValues.sessionName}`);
    }
    
    // Show success message
    ui.alert(
      'Initialization Complete',
      'YSL Hub has been successfully initialized. You can now use the system features from the YSL Hub menu.',
      ui.ButtonSet.OK
    );
    
    // Refresh menu to show operational options
    createMenu();
    
    // Log successful initialization
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('System initialization completed successfully', 'INFO', 'fullInitialization');
    }
  } catch (error) {
    // Handle initialization errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'fullInitialization', 
        'Error during system initialization. Please check the configuration and try again.');
    } else {
      Logger.log(`Initialization error: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Initialization Failed',
        `An error occurred during initialization: ${error.message}. Please check the configuration and try again.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}

/**
 * Shows the system configuration dialog for updating configuration parameters.
 */
function showConfigurationDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let assumptionsSheet = ss.getSheetByName('Assumptions');
    
    if (!assumptionsSheet) {
      // If Assumptions sheet doesn't exist, create it
      prepareAssumptionsSheet(GlobalFunctions.safeGetProperty(CONFIG.SESSION_NAME) || '');
      assumptionsSheet = ss.getSheetByName('Assumptions');
    } else {
      // Update the Assumptions sheet with current values
      const sessionName = GlobalFunctions.safeGetProperty(CONFIG.SESSION_NAME) || '';
      const rosterFolderUrl = GlobalFunctions.safeGetProperty(CONFIG.ROSTER_FOLDER_URL) || '';
      const reportTemplateUrl = GlobalFunctions.safeGetProperty(CONFIG.REPORT_TEMPLATE_URL) || '';
      const swimmerRecordsUrl = GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL) || '';
      const parentHandbookUrl = GlobalFunctions.safeGetProperty(CONFIG.PARENT_HANDBOOK_URL) || '';
      const sessionProgramsUrl = GlobalFunctions.safeGetProperty(CONFIG.SESSION_PROGRAMS_URL) || '';
      
      assumptionsSheet.getRange('B7').setValue(sessionName);
      assumptionsSheet.getRange('B8').setValue(rosterFolderUrl);
      assumptionsSheet.getRange('B9').setValue(reportTemplateUrl);
      assumptionsSheet.getRange('B10').setValue(swimmerRecordsUrl);
      assumptionsSheet.getRange('B11').setValue(parentHandbookUrl);
      assumptionsSheet.getRange('B12').setValue(sessionProgramsUrl);
    }
    
    // Activate the Assumptions sheet to show configuration
    assumptionsSheet.activate();
    
    // Add a button for applying changes
    let applyChangesButton = null;
    const drawings = assumptionsSheet.getDrawings();
    for (let i = 0; i < drawings.length; i++) {
      if (drawings[i].getAltDescription() === 'ApplyChangesButton') {
        applyChangesButton = drawings[i];
        break;
      }
    }
    
    if (!applyChangesButton) {
      // Create a text box that looks like a button
      const buttonCell = assumptionsSheet.getRange(15, 2);
      assumptionsSheet.setRowHeight(15, 30);
      
      const button = assumptionsSheet.insertTextBox('Apply Configuration Changes');
      button.setPosition(buttonCell.getRow(), buttonCell.getColumn(), 0, 0);
      button.setWidth(200);
      button.setHeight(30);
      button.getText().setFontSize(12).setFontWeight('bold');
      button.setFill('#4285F4');
      button.getText().setForegroundColor('#FFFFFF');
      button.setBorder(true, true, true, true, true, true, '#3367D6', null);
      button.setAltDescription('ApplyChangesButton');
    }
    
    ui.alert(
      'System Configuration',
      'You can update the configuration parameters in the Assumptions sheet. After making changes, ' +
      'click the "Apply Configuration Changes" button at the bottom of the sheet.\n\n' +
      'Note: Some changes may require restarting the system.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'showConfigurationDialog', 
        'Error showing configuration dialog. Please try again or contact support.');
    } else {
      Logger.log(`Configuration dialog error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error showing configuration: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Applies configuration changes from the Assumptions sheet.
 */
function applyConfigurationChanges() {
  try {
    const ui = SpreadsheetApp.getUi();
    const result = ui.alert(
      'Apply Configuration Changes',
      'Are you sure you want to apply the configuration changes? This may affect system operation.',
      ui.ButtonSet.YES_NO
    );
    
    if (result === ui.Button.YES) {
      // Extract and store updated configuration values
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const assumptionsSheet = ss.getSheetByName('Assumptions');
      
      if (!assumptionsSheet) {
        ui.alert('Error', 'Assumptions sheet not found.', ui.ButtonSet.OK);
        return;
      }
      
      // Extract updated configuration
      const configData = assumptionsSheet.getRange('A7:B12').getValues();
      const configValues = {
        sessionName: configData[0][1],
        rosterFolderUrl: configData[1][1],
        reportTemplateUrl: configData[2][1],
        swimmerRecordsUrl: configData[3][1],
        parentHandbookUrl: configData[4][1],
        sessionProgramsUrl: configData[5][1]
      };
      
      // Log the configuration changes
      if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
        ErrorHandling.logMessage('Applying configuration changes', 'INFO', 'applyConfigurationChanges');
      }
      
      // Validate and store updated configuration
      let changesMade = false;
      let instructorDataChanged = false;
      
      Object.entries(configValues).forEach(([key, value]) => {
        const propKey = CONFIG[key.toUpperCase()];
        const currentValue = GlobalFunctions.safeGetProperty(propKey);
        
        if (value && value !== currentValue) {
          GlobalFunctions.safeSetProperty(propKey, value);
          changesMade = true;
          
          // Log each change
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Configuration changed: ${key} = ${value}`, 'INFO', 'applyConfigurationChanges');
          }
          
          // Check if instructor data URL changed
          if (key === 'sessionProgramsUrl') {
            instructorDataChanged = true;
          }
        }
      });
      
      // Update instructor information if Session Programs URL changed
      if (instructorDataChanged && configValues.sessionProgramsUrl) {
        try {
          DataIntegrationModule.importInstructorData(configValues.sessionProgramsUrl);
          ui.alert(
            'Instructor Data Updated',
            'Instructor information has been updated from the Session Programs workbook.',
            ui.ButtonSet.OK
          );
        } catch (error) {
          if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
            ErrorHandling.handleError(error, 'applyConfigurationChanges', 
              'Failed to update instructor information. Other configuration changes were applied successfully.');
          } else {
            Logger.log(`Instructor data update error: ${error.message}`);
            ui.alert(
              'Instructor Update Failed',
              `Failed to update instructor information: ${error.message}`,
              ui.ButtonSet.OK
            );
          }
        }
      }
      
      if (changesMade) {
        // Record configuration update in version control
        if (VersionControl && typeof VersionControl.recordUpdate === 'function') {
          VersionControl.recordUpdate(VersionControl.getCurrentVersion(), 
            'System configuration updated by administrator');
        }
        
        ui.alert(
          'Configuration Updated',
          'The configuration changes have been applied successfully.',
          ui.ButtonSet.OK
        );
      } else {
        ui.alert(
          'No Changes',
          'No configuration changes were detected.',
          ui.ButtonSet.OK
        );
      }
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'applyConfigurationChanges', 
        'Error applying configuration changes. Please try again or contact support.');
    } else {
      Logger.log(`Configuration update error: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error applying configuration changes: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Displays information about the YSL Hub system.
 */
function showAboutDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    let versionInfo = 'Version 2.0.0 (2025-04-27)';
    
    // Get version info if available
    if (VersionControl && typeof VersionControl.getVersionInfo === 'function') {
      const vi = VersionControl.getVersionInfo();
      versionInfo = `Version ${vi.currentVersion} (${vi.releaseDate})`;
    }
    
    ui.alert(
      'About YSL Hub',
      'YSL Hub is a comprehensive management system for swim lessons, ' +
      'providing class management, assessment tracking, report generation, ' +
      'and communication capabilities.\n\n' +
      `${versionInfo}\n` +
      'PenBay YMCA Aquatics Department',
      ui.ButtonSet.OK
    );
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error showing about dialog: ${error.message}`, 'ERROR', 'showAboutDialog');
    } else {
      Logger.log(`Error showing about dialog: ${error.message}`);
    }
  }
}

/**
 * Retrieves the current system configuration.
 * @return {Object} The current configuration settings
 */
function getSystemConfiguration() {
  return {
    sessionName: GlobalFunctions.safeGetProperty(CONFIG.SESSION_NAME),
    rosterFolderUrl: GlobalFunctions.safeGetProperty(CONFIG.ROSTER_FOLDER_URL),
    reportTemplateUrl: GlobalFunctions.safeGetProperty(CONFIG.REPORT_TEMPLATE_URL),
    swimmerRecordsUrl: GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL),
    parentHandbookUrl: GlobalFunctions.safeGetProperty(CONFIG.PARENT_HANDBOOK_URL),
    sessionProgramsUrl: GlobalFunctions.safeGetProperty(CONFIG.SESSION_PROGRAMS_URL),
    isInitialized: GlobalFunctions.safeGetProperty(CONFIG.INITIALIZED) === 'true'
  };
}

function fixSystemInitializationProperty() {
  try {
    // Check which property is being used
    const scriptProps = PropertiesService.getScriptProperties();
    const systemInitialized = scriptProps.getProperty('systemInitialized');
    const configInitialized = scriptProps.getProperty('INITIALIZED');
    
    // Set both to ensure at least one works
    if (systemInitialized === 'true' || configInitialized === 'true') {
      scriptProps.setProperty('systemInitialized', 'true');
      scriptProps.setProperty('INITIALIZED', 'true');
      
      SpreadsheetApp.getUi().alert(
        'Property Fixed',
        'System initialization property has been corrected. Please reload the page to see the menu.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      scriptProps.setProperty('systemInitialized', 'true');
      scriptProps.setProperty('INITIALIZED', 'true');
      
      SpreadsheetApp.getUi().alert(
        'System Initialized',
        'System has been manually marked as initialized. Please reload the page to see the menu.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Try to refresh the menu
    createMenu();
    
    return true;
  } catch (error) {
    Logger.log(`Error fixing initialization property: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'Error',
      `Failed to fix initialization property: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}

// Make functions available to other modules
// Make functions available to other modules
const AdministrativeModule = {
  createMenu: createMenu,
  showInitializationDialog: showInitializationDialog,
  prepareAssumptionsSheet: prepareAssumptionsSheet,
  prepareAssumptionsSheetWithValues: prepareAssumptionsSheetWithValues,
  performInitialization: performInitialization,
  fullInitialization: fullInitialization,
  showConfigurationDialog: showConfigurationDialog,
  applyConfigurationChanges: applyConfigurationChanges,
  showAboutDialog: showAboutDialog,
  getSystemConfiguration: getSystemConfiguration
};