/**
 * YSL Hub v2 Menu Wrapper Functions
 * 
 * This module provides global wrapper functions for module methods
 * that need to be called from the UI menu system. These functions
 * provide direct access from the menu to the actual implementation functions.
 * 
 * @author PenBay YMCA
 * @version 2.0
 * @date 2025-04-27
 */

/**
 * Data Integration Module Functions
 */
function DataIntegrationModule_updateClassSelector() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Update Class Selector', 'INFO', 'DataIntegrationModule_updateClassSelector');
  }
  return DataIntegrationModule.updateClassSelector();
}

function DataIntegrationModule_refreshRosterData() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Refresh Roster Data', 'INFO', 'DataIntegrationModule_refreshRosterData');
  }
  return DataIntegrationModule.refreshRosterData();
}

function DataIntegrationModule_pushAssessmentsToSwimmerLog() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Push Assessments to Swimmer Log', 'INFO', 'DataIntegrationModule_pushAssessmentsToSwimmerLog');
  }
  return DataIntegrationModule.pushAssessmentsToSwimmerLog();
}

function DataIntegrationModule_pullAssessmentCriteria() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Pull Assessment Criteria', 'INFO', 'DataIntegrationModule_pullAssessmentCriteria');
  }
  return DataIntegrationModule.pullAssessmentCriteria();
}

function DataIntegrationModule_updateInstructorData() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Update Instructor Data', 'INFO', 'DataIntegrationModule_updateInstructorData');
  }
  
  try {
    const config = AdministrativeModule.getSystemConfiguration();
    if (!config.sessionProgramsUrl) {
      SpreadsheetApp.getUi().alert(
        'Missing Configuration',
        'Session Programs workbook URL is not configured. Please update system configuration first.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Call the function from the DataIntegrationModule
    const result = DataIntegrationModule.importInstructorData(config.sessionProgramsUrl);
    
    // Show success message if the instructor data was imported
    SpreadsheetApp.getUi().alert(
      'Instructor Data Update',
      'Instructor information has been updated from the Session Programs workbook.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return result;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'DataIntegrationModule_updateInstructorData', 
        'Failed to update instructor information. Please check your configuration and try again.');
    } else {
      Logger.log(`Instructor data update error: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Update Failed',
        `Failed to update instructor information: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Instructor Resource Module Functions
 */
function InstructorResourceModule_generateInstructorSheets() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Generate Instructor Sheets', 'INFO', 'InstructorResourceModule_generateInstructorSheets');
  }
  return InstructorResourceModule.generateInstructorSheets();
}

/**
 * Reporting Module Functions
 */
function ReportingModule_generateMidSessionReports() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Generate Mid-Session Reports', 'INFO', 'ReportingModule_generateMidSessionReports');
  }
  return ReportingModule.generateMidSessionReports();
}

function ReportingModule_generateEndSessionReports() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Generate End-Session Reports', 'INFO', 'ReportingModule_generateEndSessionReports');
  }
  return ReportingModule.generateEndSessionReports();
}

/**
 * Communication Module Functions
 */
function CommunicationModule_emailClassParticipants() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Email Class Participants', 'INFO', 'CommunicationModule_emailClassParticipants');
  }
  return CommunicationModule.emailClassParticipants();
}

function CommunicationModule_sendClassAnnouncements() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Send Class Announcements', 'INFO', 'CommunicationModule_sendClassAnnouncements');
  }
  return CommunicationModule.sendClassAnnouncements();
}

function CommunicationModule_sendReadyAnnouncements() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Send Ready Announcements', 'INFO', 'CommunicationModule_sendReadyAnnouncements');
  }
  return CommunicationModule.sendReadyAnnouncements();
}

/**
 * Administrative Module Functions
 */
function AdministrativeModule_fullInitialization() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Full Initialization', 'INFO', 'AdministrativeModule_fullInitialization');
  }
  return AdministrativeModule.fullInitialization();
}

function AdministrativeModule_showInitializationDialog() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Show Initialization Dialog', 'INFO', 'AdministrativeModule_showInitializationDialog');
  }
  return AdministrativeModule.showInitializationDialog();
}

function AdministrativeModule_showConfigurationDialog() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Show Configuration Dialog', 'INFO', 'AdministrativeModule_showConfigurationDialog');
  }
  return AdministrativeModule.showConfigurationDialog();
}

function AdministrativeModule_applyConfigurationChanges() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Apply Configuration Changes', 'INFO', 'AdministrativeModule_applyConfigurationChanges');
  }
  return AdministrativeModule.applyConfigurationChanges();
}

function AdministrativeModule_showAboutDialog() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Show About Dialog', 'INFO', 'AdministrativeModule_showAboutDialog');
  }
  return AdministrativeModule.showAboutDialog();
}

/**
 * Version Control Module Functions
 */
function VersionControl_showDiagnostics() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Show Diagnostics', 'INFO', 'VersionControl_showDiagnostics');
  }
  return VersionControl.showDiagnostics();
}

function VersionControl_showVersionInfo() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Show Version Info', 'INFO', 'VersionControl_showVersionInfo');
  }
  return VersionControl.showVersionInfo();
}

/**
 * Error Handling Module Functions
 */
function ErrorHandling_showLogViewer() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Show Log Viewer', 'INFO', 'ErrorHandling_showLogViewer');
  }
  return ErrorHandling.showLogViewer();
}

function ErrorHandling_hideLogSheet() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Hide Log Sheet', 'INFO', 'ErrorHandling_hideLogSheet');
  }
  return ErrorHandling.hideLogSheet();
}

function ErrorHandling_clearLog() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Clear Log', 'INFO', 'ErrorHandling_clearLog');
  }
  return ErrorHandling.clearLog();
}

function ErrorHandling_exportLog() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Export Log', 'INFO', 'ErrorHandling_exportLog');
  }
  return ErrorHandling.exportLog();
}

function DataIntegrationModule_reportAssessmentCriteria() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Report Assessment Criteria', 'INFO', 'DataIntegrationModule_reportAssessmentCriteria');
  }
  return DataIntegrationModule.reportAssessmentCriteria();
}

function DataIntegrationModule_diagnoseCriteriaImport() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Diagnose Criteria Import', 'INFO', 'DataIntegrationModule_diagnoseCriteriaImport');
  }
  return DataIntegrationModule.diagnoseCriteriaImport();
}

function CommunicationModule_sendWelcomeEmails() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Send Welcome Emails', 'INFO', 'CommunicationModule_sendWelcomeEmails');
  }
  return CommunicationModule.sendWelcomeEmails();
}

function CommunicationModule_testWelcomeEmail() {
  if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage('Menu: Test Welcome Email', 'INFO', 'CommunicationModule_testWelcomeEmail');
  }
  return CommunicationModule.testWelcomeEmail();
}