/**
 * YSL Hub v2 Session Transition Module
 * 
 * This module handles the process of transitioning the system from one session
 * to another. It provides functions to archive old session data, update
 * configuration for a new session, and prepare the system for a new group of
 * swim classes.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-16
 */

// Session transition state properties
const TRANSITION_PROPS = {
  IN_PROGRESS: 'transitionInProgress',
  SOURCE_SESSION: 'transitionSourceSession',
  TARGET_SESSION: 'transitionTargetSession',
  START_DATE: 'transitionStartDate',
  CURRENT_STEP: 'transitionCurrentStep',
  COMPLETED_STEPS: 'transitionCompletedSteps',
  ARCHIVED_URL: 'transitionArchivedUrl'
};

// Session transition steps
const TRANSITION_STEPS = {
  INITIALIZE: 'initialize',
  EXPORT_DATA: 'exportData',
  ARCHIVE_WORKBOOK: 'archiveWorkbook',
  UPDATE_CONFIG: 'updateConfig',
  CLEAR_DATA: 'clearData',
  IMPORT_NEW_DATA: 'importNewData',
  FINALIZE: 'finalize'
};

/**
 * Starts a new session transition process
 * 
 * @returns Success status
 */
function startSessionTransition() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting session transition', 'INFO', 'startSessionTransition');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Check if a transition is already in progress
    const scriptProps = PropertiesService.getScriptProperties();
    const inProgress = scriptProps.getProperty(TRANSITION_PROPS.IN_PROGRESS) === 'true';
    
    if (inProgress) {
      const result = ui.alert(
        'Transition In Progress',
        'A session transition is already in progress. Would you like to resume it?',
        ui.ButtonSet.YES_NO
      );
      
      if (result === ui.Button.YES) {
        return resumeSessionTransition();
      } else {
        return false;
      }
    }
    
    // Get current session info
    const config = AdministrativeModule.getSystemConfiguration();
    const currentSession = config.sessionName || 'Current Session';
    
    // Ask for new session name
    const sessionResult = ui.prompt(
      'New Session',
      `You are about to transition from "${currentSession}" to a new session.\n\nPlease enter the name for the new session (e.g., "Summer2 2025"):`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (sessionResult.getSelectedButton() !== ui.Button.OK) {
      return false;
    }
    
    const newSession = sessionResult.getResponseText().trim();
    
    if (!newSession) {
      ui.alert(
        'Error',
        'Session name cannot be empty. Please try again.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Confirm the transition
    const confirmResult = ui.alert(
      'Confirm Transition',
      `You are about to transition from "${currentSession}" to "${newSession}".\n\n` +
      'This process will:\n' +
      '1. Archive current session data\n' +
      '2. Create a backup of this workbook\n' +
      '3. Update configuration for the new session\n' +
      '4. Clear assessment data while preserving system sheets\n' +
      '5. Prepare the system for the new session\n\n' +
      'Do you want to continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResult !== ui.Button.YES) {
      return false;
    }
    
    // Set transition properties
    scriptProps.setProperty(TRANSITION_PROPS.IN_PROGRESS, 'true');
    scriptProps.setProperty(TRANSITION_PROPS.SOURCE_SESSION, currentSession);
    scriptProps.setProperty(TRANSITION_PROPS.TARGET_SESSION, newSession);
    scriptProps.setProperty(TRANSITION_PROPS.START_DATE, new Date().toISOString());
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, TRANSITION_STEPS.INITIALIZE);
    scriptProps.setProperty(TRANSITION_PROPS.COMPLETED_STEPS, JSON.stringify([]));
    
    // Initialize the transition
    const initResult = initializeTransition(currentSession, newSession);
    
    if (!initResult) {
      scriptProps.deleteProperty(TRANSITION_PROPS.IN_PROGRESS);
      return false;
    }
    
    // Proceed to next step - Export Data
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, TRANSITION_STEPS.EXPORT_DATA);
    const exportResult = exportSessionData(currentSession);
    
    if (!exportResult) {
      ui.alert(
        'Transition Paused',
        'The session transition has been paused at the data export step. You can resume it later using "Resume Session Transition" from the System menu.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Update completed steps
    updateCompletedSteps(TRANSITION_STEPS.EXPORT_DATA);
    
    // Proceed to next step - Archive Workbook
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, TRANSITION_STEPS.ARCHIVE_WORKBOOK);
    const archiveResult = archiveWorkbook(currentSession);
    
    if (!archiveResult) {
      ui.alert(
        'Transition Paused',
        'The session transition has been paused at the workbook archive step. You can resume it later using "Resume Session Transition" from the System menu.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Update completed steps
    updateCompletedSteps(TRANSITION_STEPS.ARCHIVE_WORKBOOK);
    
    // Proceed to next step - Update Configuration
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, TRANSITION_STEPS.UPDATE_CONFIG);
    const configResult = updateConfiguration(newSession);
    
    if (!configResult) {
      ui.alert(
        'Transition Paused',
        'The session transition has been paused at the configuration update step. You can resume it later using "Resume Session Transition" from the System menu.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Update completed steps
    updateCompletedSteps(TRANSITION_STEPS.UPDATE_CONFIG);
    
    // Proceed to next step - Clear Data
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, TRANSITION_STEPS.CLEAR_DATA);
    const clearResult = clearSessionData();
    
    if (!clearResult) {
      ui.alert(
        'Transition Paused',
        'The session transition has been paused at the data clearing step. You can resume it later using "Resume Session Transition" from the System menu.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Update completed steps
    updateCompletedSteps(TRANSITION_STEPS.CLEAR_DATA);
    
    // Proceed to next step - Import New Data
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, TRANSITION_STEPS.IMPORT_NEW_DATA);
    const importResult = importNewSessionData(newSession);
    
    if (!importResult) {
      ui.alert(
        'Transition Paused',
        'The session transition has been paused at the data import step. You can resume it later using "Resume Session Transition" from the System menu.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Update completed steps
    updateCompletedSteps(TRANSITION_STEPS.IMPORT_NEW_DATA);
    
    // Proceed to final step - Finalize Transition
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, TRANSITION_STEPS.FINALIZE);
    const finalizeResult = finalizeTransition(newSession);
    
    if (!finalizeResult) {
      ui.alert(
        'Transition Paused',
        'The session transition has been paused at the finalization step. You can resume it later using "Resume Session Transition" from the System menu.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Update completed steps
    updateCompletedSteps(TRANSITION_STEPS.FINALIZE);
    
    // Clear transition properties
    clearTransitionProperties();
    
    // Show success message
    ui.alert(
      'Transition Complete',
      `The session transition from "${currentSession}" to "${newSession}" has been completed successfully.\n\n` +
      'The system is now ready for the new session.',
      ui.ButtonSet.OK
    );
    
    // Record in history
    if (HistoryModule && typeof HistoryModule.recordHistoryEvent === 'function') {
      HistoryModule.recordHistoryEvent(
        HistoryModule.EVENT_TYPES.SESSION_TRANSITION,
        `Transitioned from "${currentSession}" to "${newSession}"`,
        'Complete session transition process'
      );
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'startSessionTransition', 
        'Error starting session transition. Please try again or contact support.');
    } else {
      Logger.log(`Error starting session transition: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to start session transition: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Resumes an in-progress session transition
 * 
 * @returns Success status
 */
function resumeSessionTransition() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Resuming session transition', 'INFO', 'resumeSessionTransition');
    }
    
    const ui = SpreadsheetApp.getUi();
    
    // Check if a transition is in progress
    const scriptProps = PropertiesService.getScriptProperties();
    const inProgress = scriptProps.getProperty(TRANSITION_PROPS.IN_PROGRESS) === 'true';
    
    if (!inProgress) {
      ui.alert(
        'No Transition in Progress',
        'There is no session transition in progress to resume. Please start a new transition.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Get transition properties
    const sourceSession = scriptProps.getProperty(TRANSITION_PROPS.SOURCE_SESSION);
    const targetSession = scriptProps.getProperty(TRANSITION_PROPS.TARGET_SESSION);
    const currentStep = scriptProps.getProperty(TRANSITION_PROPS.CURRENT_STEP);
    
    // Show confirmation
    const confirmResult = ui.alert(
      'Resume Transition',
      `Resuming session transition from "${sourceSession}" to "${targetSession}".\n\n` +
      `Current step: ${getStepDescription(currentStep)}\n\n` +
      'Do you want to continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResult !== ui.Button.YES) {
      return false;
    }
    
    // Resume from current step
    switch (currentStep) {
      case TRANSITION_STEPS.INITIALIZE:
        return initializeTransition(sourceSession, targetSession) && 
               continueTransition(TRANSITION_STEPS.EXPORT_DATA);
        
      case TRANSITION_STEPS.EXPORT_DATA:
        return exportSessionData(sourceSession) && 
               continueTransition(TRANSITION_STEPS.ARCHIVE_WORKBOOK);
        
      case TRANSITION_STEPS.ARCHIVE_WORKBOOK:
        return archiveWorkbook(sourceSession) && 
               continueTransition(TRANSITION_STEPS.UPDATE_CONFIG);
        
      case TRANSITION_STEPS.UPDATE_CONFIG:
        return updateConfiguration(targetSession) && 
               continueTransition(TRANSITION_STEPS.CLEAR_DATA);
        
      case TRANSITION_STEPS.CLEAR_DATA:
        return clearSessionData() && 
               continueTransition(TRANSITION_STEPS.IMPORT_NEW_DATA);
        
      case TRANSITION_STEPS.IMPORT_NEW_DATA:
        return importNewSessionData(targetSession) && 
               continueTransition(TRANSITION_STEPS.FINALIZE);
        
      case TRANSITION_STEPS.FINALIZE:
        return finalizeTransition(targetSession);
        
      default:
        // Unknown step - reset and start over
        ui.alert(
          'Unknown Step',
          'The transition process is in an unknown state. Please start a new transition.',
          ui.ButtonSet.OK
        );
        clearTransitionProperties();
        return false;
    }
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'resumeSessionTransition', 
        'Error resuming session transition. Please try again or contact support.');
    } else {
      Logger.log(`Error resuming session transition: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to resume session transition: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Gets a human-readable description of a transition step
 * 
 * @param step - The transition step
 * @returns Description of the step
 */
function getStepDescription(step) {
  switch (step) {
    case TRANSITION_STEPS.INITIALIZE:
      return 'Initialize Transition';
    case TRANSITION_STEPS.EXPORT_DATA:
      return 'Export Current Session Data';
    case TRANSITION_STEPS.ARCHIVE_WORKBOOK:
      return 'Archive Current Workbook';
    case TRANSITION_STEPS.UPDATE_CONFIG:
      return 'Update System Configuration';
    case TRANSITION_STEPS.CLEAR_DATA:
      return 'Clear Current Session Data';
    case TRANSITION_STEPS.IMPORT_NEW_DATA:
      return 'Import New Session Data';
    case TRANSITION_STEPS.FINALIZE:
      return 'Finalize Transition';
    default:
      return 'Unknown Step';
  }
}

/**
 * Continues the transition process from a specific step
 * 
 * @param nextStep - The next step to execute
 * @returns Success status
 */
function continueTransition(nextStep) {
  try {
    // Update current step
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty(TRANSITION_PROPS.CURRENT_STEP, nextStep);
    
    // Get session info
    const sourceSession = scriptProps.getProperty(TRANSITION_PROPS.SOURCE_SESSION);
    const targetSession = scriptProps.getProperty(TRANSITION_PROPS.TARGET_SESSION);
    
    // Execute the step
    switch (nextStep) {
      case TRANSITION_STEPS.EXPORT_DATA:
        // Update completed steps for previous step
        updateCompletedSteps(TRANSITION_STEPS.INITIALIZE);
        return exportSessionData(sourceSession) && 
               continueTransition(TRANSITION_STEPS.ARCHIVE_WORKBOOK);
        
      case TRANSITION_STEPS.ARCHIVE_WORKBOOK:
        // Update completed steps for previous step
        updateCompletedSteps(TRANSITION_STEPS.EXPORT_DATA);
        return archiveWorkbook(sourceSession) && 
               continueTransition(TRANSITION_STEPS.UPDATE_CONFIG);
        
      case TRANSITION_STEPS.UPDATE_CONFIG:
        // Update completed steps for previous step
        updateCompletedSteps(TRANSITION_STEPS.ARCHIVE_WORKBOOK);
        return updateConfiguration(targetSession) && 
               continueTransition(TRANSITION_STEPS.CLEAR_DATA);
        
      case TRANSITION_STEPS.CLEAR_DATA:
        // Update completed steps for previous step
        updateCompletedSteps(TRANSITION_STEPS.UPDATE_CONFIG);
        return clearSessionData() && 
               continueTransition(TRANSITION_STEPS.IMPORT_NEW_DATA);
        
      case TRANSITION_STEPS.IMPORT_NEW_DATA:
        // Update completed steps for previous step
        updateCompletedSteps(TRANSITION_STEPS.CLEAR_DATA);
        return importNewSessionData(targetSession) && 
               continueTransition(TRANSITION_STEPS.FINALIZE);
        
      case TRANSITION_STEPS.FINALIZE:
        // Update completed steps for previous step
        updateCompletedSteps(TRANSITION_STEPS.IMPORT_NEW_DATA);
        const result = finalizeTransition(targetSession);
        
        if (result) {
          updateCompletedSteps(TRANSITION_STEPS.FINALIZE);
          clearTransitionProperties();
          
          // Show success message
          SpreadsheetApp.getUi().alert(
            'Transition Complete',
            `The session transition from "${sourceSession}" to "${targetSession}" has been completed successfully.\n\n` +
            'The system is now ready for the new session.',
            SpreadsheetApp.getUi().ButtonSet.OK
          );
          
          // Record in history
          if (HistoryModule && typeof HistoryModule.recordHistoryEvent === 'function') {
            HistoryModule.recordHistoryEvent(
              HistoryModule.EVENT_TYPES.SESSION_TRANSITION,
              `Transitioned from "${sourceSession}" to "${targetSession}"`,
              'Complete session transition process'
            );
          }
        }
        
        return result;
        
      default:
        // Unknown step - reset and start over
        SpreadsheetApp.getUi().alert(
          'Unknown Step',
          'The transition process is in an unknown state. Please start a new transition.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );
        clearTransitionProperties();
        return false;
    }
  } catch (error) {
    Logger.log(`Error continuing transition: ${error.message}`);
    return false;
  }
}

/**
 * Updates the list of completed transition steps
 * 
 * @param completedStep - The step that was just completed
 */
function updateCompletedSteps(completedStep) {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    const completedStepsJson = scriptProps.getProperty(TRANSITION_PROPS.COMPLETED_STEPS) || '[]';
    
    let completedSteps = [];
    try {
      completedSteps = JSON.parse(completedStepsJson);
    } catch (e) {
      completedSteps = [];
    }
    
    // Add the step if not already in the list
    if (!completedSteps.includes(completedStep)) {
      completedSteps.push(completedStep);
    }
    
    // Save updated list
    scriptProps.setProperty(TRANSITION_PROPS.COMPLETED_STEPS, JSON.stringify(completedSteps));
  } catch (error) {
    Logger.log(`Error updating completed steps: ${error.message}`);
  }
}

/**
 * Clears all transition properties
 */
function clearTransitionProperties() {
  try {
    const scriptProps = PropertiesService.getScriptProperties();
    
    scriptProps.deleteProperty(TRANSITION_PROPS.IN_PROGRESS);
    scriptProps.deleteProperty(TRANSITION_PROPS.SOURCE_SESSION);
    scriptProps.deleteProperty(TRANSITION_PROPS.TARGET_SESSION);
    scriptProps.deleteProperty(TRANSITION_PROPS.START_DATE);
    scriptProps.deleteProperty(TRANSITION_PROPS.CURRENT_STEP);
    scriptProps.deleteProperty(TRANSITION_PROPS.COMPLETED_STEPS);
    // Don't delete ARCHIVED_URL as it may be needed for reference
  } catch (error) {
    Logger.log(`Error clearing transition properties: ${error.message}`);
  }
}

/**
 * Initializes the transition process
 * 
 * @param sourceSession - The source session name
 * @param targetSession - The target session name
 * @returns Success status
 */
function initializeTransition(sourceSession, targetSession) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Initializing session transition', 'INFO', 'initializeTransition');
    }
    
    // Create transition sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let transitionSheet = ss.getSheetByName('TransitionWorkspace');
    
    if (transitionSheet) {
      // Clear existing sheet
      transitionSheet.clear();
    } else {
      // Create new sheet
      transitionSheet = ss.insertSheet('TransitionWorkspace');
    }
    
    // Set up sheet header
    transitionSheet.getRange('A1:E1').merge()
      .setValue(`Session Transition: ${sourceSession} → ${targetSession}`)
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Add transition info
    transitionSheet.getRange('A3').setValue('Source Session:');
    transitionSheet.getRange('B3').setValue(sourceSession);
    transitionSheet.getRange('A4').setValue('Target Session:');
    transitionSheet.getRange('B4').setValue(targetSession);
    transitionSheet.getRange('A5').setValue('Start Date:');
    transitionSheet.getRange('B5').setValue(new Date());
    
    // Add status section
    transitionSheet.getRange('A7:E7').merge()
      .setValue('Transition Status')
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Add step statuses
    let currentRow = 8;
    Object.entries(TRANSITION_STEPS).forEach(([key, step]) => {
      transitionSheet.getRange(currentRow, 1).setValue(getStepDescription(step));
      transitionSheet.getRange(currentRow, 2).setValue('Pending');
      currentRow++;
    });
    
    // Format sheet
    transitionSheet.setColumnWidth(1, 200);
    transitionSheet.setColumnWidth(2, 150);
    transitionSheet.setColumnWidth(3, 150);
    transitionSheet.setColumnWidth(4, 150);
    transitionSheet.setColumnWidth(5, 150);
    
    // Check for required configuration
    const config = AdministrativeModule.getSystemConfiguration();
    
    if (!config.swimmerRecordsUrl) {
      SpreadsheetApp.getUi().alert(
        'Missing Configuration',
        'Swimmer Records URL is not configured. Please update system configuration before continuing.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    if (!config.reportTemplateUrl) {
      SpreadsheetApp.getUi().alert(
        'Missing Configuration',
        'Report Template URL is not configured. Please update system configuration before continuing.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return false;
    }
    
    // Show the transition sheet
    transitionSheet.activate();
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.INITIALIZE, 'Completed');
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'initializeTransition', 
        'Error initializing transition. Please try again or contact support.');
    } else {
      Logger.log(`Error initializing transition: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to initialize transition: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    return false;
  }
}

/**
 * Updates the status of a transition step
 * 
 * @param step - The transition step
 * @param status - The new status
 */
function updateTransitionStatus(step, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const transitionSheet = ss.getSheetByName('TransitionWorkspace');
    
    if (!transitionSheet) {
      return;
    }
    
    // Find the step row
    const dataRange = transitionSheet.getDataRange();
    const data = dataRange.getValues();
    
    for (let i = 7; i < data.length; i++) {
      if (data[i][0] === getStepDescription(step)) {
        // Update status
        transitionSheet.getRange(i + 1, 2).setValue(status);
        
        // Add timestamp for completed steps
        if (status === 'Completed') {
          transitionSheet.getRange(i + 1, 3).setValue(new Date());
        }
        
        break;
      }
    }
  } catch (error) {
    Logger.log(`Error updating transition status: ${error.message}`);
  }
}

/**
 * Exports current session data to archive sheets
 * 
 * @param sourceSession - The source session name
 * @returns Success status
 */
function exportSessionData(sourceSession) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Exporting session data', 'INFO', 'exportSessionData');
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.EXPORT_DATA, 'In Progress');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create archive sheets for classes and roster
    const archivePrefix = `Archive_${sourceSession}_`;
    
    // Archive Classes sheet
    const classesSheet = ss.getSheetByName('Classes');
    if (classesSheet) {
      const archiveClassesName = archivePrefix + 'Classes';
      let archiveSheet = ss.getSheetByName(archiveClassesName);
      
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveClassesName);
      } else {
        archiveSheet.clear();
      }
      
      // Copy data
      const classesData = classesSheet.getDataRange();
      classesData.copyTo(archiveSheet.getRange(1, 1));
      
      // Format sheet
      archiveSheet.hideSheet();
    }
    
    // Archive Roster sheet
    const rosterSheet = ss.getSheetByName('Roster');
    if (rosterSheet) {
      const archiveRosterName = archivePrefix + 'Roster';
      let archiveSheet = ss.getSheetByName(archiveRosterName);
      
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveRosterName);
      } else {
        archiveSheet.clear();
      }
      
      // Copy data
      const rosterData = rosterSheet.getDataRange();
      rosterData.copyTo(archiveSheet.getRange(1, 1));
      
      // Format sheet
      archiveSheet.hideSheet();
    }
    
    // Archive Group Lesson Tracker if it exists
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    if (trackerSheet) {
      const archiveTrackerName = archivePrefix + 'GroupLessonTracker';
      let archiveSheet = ss.getSheetByName(archiveTrackerName);
      
      if (!archiveSheet) {
        archiveSheet = ss.insertSheet(archiveTrackerName);
      } else {
        archiveSheet.clear();
      }
      
      // Copy data
      const trackerData = trackerSheet.getDataRange();
      trackerData.copyTo(archiveSheet.getRange(1, 1));
      
      // Format sheet
      archiveSheet.hideSheet();
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.EXPORT_DATA, 'Completed');
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'exportSessionData', 
        'Error exporting session data. Please try again or contact support.');
    } else {
      Logger.log(`Error exporting session data: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to export session data: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.EXPORT_DATA, 'Failed');
    
    return false;
  }
}

/**
 * Archives the current workbook as a backup
 * 
 * @param sourceSession - The source session name
 * @returns Success status
 */
function archiveWorkbook(sourceSession) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Archiving workbook', 'INFO', 'archiveWorkbook');
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.ARCHIVE_WORKBOOK, 'In Progress');
    
    // Get current spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ssId = ss.getId();
    const ssName = ss.getName();
    
    // Create archived copy
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const archivedName = `${ssName} - ${sourceSession} Archive (${timestamp})`;
    
    const archivedFile = DriveApp.getFileById(ssId).makeCopy(archivedName);
    const archivedUrl = archivedFile.getUrl();
    
    // Store archived URL in properties
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty(TRANSITION_PROPS.ARCHIVED_URL, archivedUrl);
    
    // Add info to transition sheet
    const transitionSheet = ss.getSheetByName('TransitionWorkspace');
    if (transitionSheet) {
      transitionSheet.getRange('A15').setValue('Archive URL:');
      transitionSheet.getRange('B15:E15').merge().setValue(archivedUrl);
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.ARCHIVE_WORKBOOK, 'Completed');
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'archiveWorkbook', 
        'Error archiving workbook. Please try again or contact support.');
    } else {
      Logger.log(`Error archiving workbook: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to archive workbook: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.ARCHIVE_WORKBOOK, 'Failed');
    
    return false;
  }
}

/**
 * Updates the system configuration for the new session
 * 
 * @param targetSession - The target session name
 * @returns Success status
 */
function updateConfiguration(targetSession) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Updating configuration', 'INFO', 'updateConfiguration');
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.UPDATE_CONFIG, 'In Progress');
    
    // Update session name in properties
    const scriptProps = PropertiesService.getScriptProperties();
    scriptProps.setProperty('sessionName', targetSession);
    scriptProps.setProperty('SESSION_NAME', targetSession);
    
    // Update Assumptions sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assumptionsSheet = ss.getSheetByName('Assumptions');
    
    if (assumptionsSheet) {
      // Find session name cell (typically B7)
      const dataRange = assumptionsSheet.getDataRange();
      const data = dataRange.getValues();
      
      for (let i = 0; i < data.length; i++) {
        if (data[i][0] === 'Session Name') {
          assumptionsSheet.getRange(i + 1, 2).setValue(targetSession);
          break;
        }
      }
    } else {
      // Create Assumptions sheet if it doesn't exist
      if (AdministrativeModule && typeof AdministrativeModule.prepareAssumptionsSheet === 'function') {
        AdministrativeModule.prepareAssumptionsSheet(targetSession);
      }
    }
    
    // Ask for roster folder URL
    const ui = SpreadsheetApp.getUi();
    const folderResult = ui.prompt(
      'Roster Folder',
      'Enter the URL of the folder containing the roster file for the new session (or leave blank to keep current setting):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (folderResult.getSelectedButton() !== ui.Button.OK) {
      // Update status
      updateTransitionStatus(TRANSITION_STEPS.UPDATE_CONFIG, 'Failed');
      return false;
    }
    
    const folderUrl = folderResult.getResponseText().trim();
    
    if (folderUrl) {
      // Update roster folder URL
      scriptProps.setProperty('rosterFolderUrl', folderUrl);
      scriptProps.setProperty('ROSTER_FOLDER_URL', folderUrl);
      
      // Update in Assumptions sheet
      if (assumptionsSheet) {
        const dataRange = assumptionsSheet.getDataRange();
        const data = dataRange.getValues();
        
        for (let i = 0; i < data.length; i++) {
          if (data[i][0] === 'Session Roster Folder URL') {
            assumptionsSheet.getRange(i + 1, 2).setValue(folderUrl);
            break;
          }
        }
      }
    }
    
    // Record change in history
    if (HistoryModule && typeof HistoryModule.recordHistoryEvent === 'function') {
      HistoryModule.recordHistoryEvent(
        HistoryModule.EVENT_TYPES.CONFIGURATION,
        `Updated configuration for new session: ${targetSession}`,
        'Session transition configuration update'
      );
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.UPDATE_CONFIG, 'Completed');
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'updateConfiguration', 
        'Error updating configuration. Please try again or contact support.');
    } else {
      Logger.log(`Error updating configuration: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to update configuration: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.UPDATE_CONFIG, 'Failed');
    
    return false;
  }
}

/**
 * Clears current session data to prepare for the new session
 * 
 * @returns Success status
 */
function clearSessionData() {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Clearing session data', 'INFO', 'clearSessionData');
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.CLEAR_DATA, 'In Progress');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Clear Classes sheet data (preserve header row)
    const classesSheet = ss.getSheetByName('Classes');
    if (classesSheet && classesSheet.getLastRow() > 1) {
      classesSheet.deleteRows(2, classesSheet.getLastRow() - 1);
    }
    
    // Clear Roster sheet data (preserve header row)
    const rosterSheet = ss.getSheetByName('Roster');
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      rosterSheet.deleteRows(2, rosterSheet.getLastRow() - 1);
    }
    
    // Remove Group Lesson Tracker sheet
    const trackerSheet = ss.getSheetByName('Group Lesson Tracker');
    if (trackerSheet) {
      ss.deleteSheet(trackerSheet);
    }
    
    // Clear CommunicationLog sheet data if it exists (preserve header row)
    const commLogSheet = ss.getSheetByName('CommunicationLog');
    if (commLogSheet && commLogSheet.getLastRow() > 1) {
      commLogSheet.deleteRows(2, commLogSheet.getLastRow() - 1);
    }
    
    // Clear any report summary sheets
    const midSessionSummary = ss.getSheetByName('Mid-Session Progress Report - Summary');
    if (midSessionSummary) {
      ss.deleteSheet(midSessionSummary);
    }
    
    const endSessionSummary = ss.getSheetByName('End-Session Assessment Report - Summary');
    if (endSessionSummary) {
      ss.deleteSheet(endSessionSummary);
    }
    
    // Clear ClassSelector sheet if it exists
    const selectorSheet = ss.getSheetByName('ClassSelector');
    if (selectorSheet) {
      ss.deleteSheet(selectorSheet);
    }
    
    // Clear TempClassSelection sheet if it exists
    const tempSelectSheet = ss.getSheetByName('TempClassSelection');
    if (tempSelectSheet) {
      ss.deleteSheet(tempSelectSheet);
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.CLEAR_DATA, 'Completed');
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'clearSessionData', 
        'Error clearing session data. Please try again or contact support.');
    } else {
      Logger.log(`Error clearing session data: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to clear session data: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.CLEAR_DATA, 'Failed');
    
    return false;
  }
}

/**
 * Imports data for the new session
 * 
 * @param targetSession - The target session name
 * @returns Success status
 */
function importNewSessionData(targetSession) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Importing new session data', 'INFO', 'importNewSessionData');
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.IMPORT_NEW_DATA, 'In Progress');
    
    const ui = SpreadsheetApp.getUi();
    const config = AdministrativeModule.getSystemConfiguration();
    
    // Get roster folder URL
    const rosterFolderUrl = config.rosterFolderUrl;
    
    if (!rosterFolderUrl) {
      ui.alert(
        'Missing Configuration',
        'Roster folder URL is not configured. Please update system configuration before continuing.',
        ui.ButtonSet.OK
      );
      
      // Update status
      updateTransitionStatus(TRANSITION_STEPS.IMPORT_NEW_DATA, 'Failed');
      
      return false;
    }
    
    // Ask if user wants to import roster data now
    const importResult = ui.alert(
      'Import Roster Data',
      `Do you want to import roster data for the new session "${targetSession}" now?\n\n` +
      'Note: This requires that a roster file exists in the configured folder with the naming convention:\n' +
      `"YSL ${targetSession} Roster"`,
      ui.ButtonSet.YES_NO
    );
    
    if (importResult === ui.Button.YES) {
      // Try to import roster data
      const rosterResult = DataIntegrationModule.refreshRosterData();
      
      if (!rosterResult) {
        ui.alert(
          'Import Warning',
          'Could not import roster data. You can try again later using "Refresh Roster Data" from the main menu.',
          ui.ButtonSet.OK
        );
      }
    }
    
    // Check if we need to update assessment criteria
    const criteriaResult = ui.alert(
      'Update Assessment Criteria',
      'Do you want to update assessment criteria from the Swimmer Records workbook?',
      ui.ButtonSet.YES_NO
    );
    
    if (criteriaResult === ui.Button.YES) {
      // Try to import criteria
      const config: any = AdministrativeModule && typeof AdministrativeModule.getSystemConfiguration === 'function' ?
        AdministrativeModule.getSystemConfiguration() : {};
      DataIntegrationModule.pullAssessmentCriteria(config.swimmerRecordsUrl || '');
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.IMPORT_NEW_DATA, 'Completed');
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'importNewSessionData', 
        'Error importing new session data. Please try again or contact support.');
    } else {
      Logger.log(`Error importing new session data: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to import new session data: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.IMPORT_NEW_DATA, 'Failed');
    
    return false;
  }
}

/**
 * Finalizes the transition process
 * 
 * @param targetSession - The target session name
 * @returns Success status
 */
function finalizeTransition(targetSession) {
  try {
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Finalizing transition', 'INFO', 'finalizeTransition');
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.FINALIZE, 'In Progress');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create a welcome message sheet
    let welcomeSheet = ss.getSheetByName('WelcomeToNewSession');
    
    if (welcomeSheet) {
      welcomeSheet.clear();
    } else {
      welcomeSheet = ss.insertSheet('WelcomeToNewSession');
    }
    
    // Retrieve archived URL
    const scriptProps = PropertiesService.getScriptProperties();
    const archivedUrl = scriptProps.getProperty(TRANSITION_PROPS.ARCHIVED_URL) || 'Unknown';
    
    // Set up welcome sheet
    welcomeSheet.getRange('A1:E1').merge()
      .setValue(`Welcome to YSL Hub - ${targetSession}`)
      .setFontWeight('bold')
      .setFontSize(16)
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    welcomeSheet.getRange('A3:E10').merge()
      .setValue(`The system has been successfully transitioned to the new session: ${targetSession}\n\n` +
                'Key information:\n\n' +
                `1. Archive of previous session is available at: ${archivedUrl}\n\n` +
                '2. To continue setting up this session:\n' +
                '   - Use "Refresh Class List" to populate the class dropdown\n' +
                '   - Generate the Group Lesson Tracker for your classes\n' +
                '   - Update any configuration as needed in System Settings\n\n' +
                '3. You can safely delete this sheet once you\'ve noted the archive URL')
      .setWrap(true);
    
    // Format sheet
    welcomeSheet.setColumnWidth(1, 30);
    for (let i = 2; i <= 4; i++) {
      welcomeSheet.setColumnWidth(i, 150);
    }
    welcomeSheet.setColumnWidth(5, 30);
    
    // Hide the transition workspace
    const transitionSheet = ss.getSheetByName('TransitionWorkspace');
    if (transitionSheet) {
      transitionSheet.hideSheet();
    }
    
    // Activate welcome sheet
    welcomeSheet.activate();
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.FINALIZE, 'Completed');
    
    // Record session transition
    const sessionHistory = scriptProps.getProperty('sessionHistory') || '[]';
    let historyArray = [];
    
    try {
      historyArray = JSON.parse(sessionHistory);
    } catch (e) {
      historyArray = [];
    }
    
    // Add new session to history
    historyArray.push({
      name: targetSession,
      date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      time: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HH:mm:ss'),
      user: Session.getEffectiveUser().getEmail(),
      details: `Transitioned from previous session. Archive URL: ${archivedUrl}`
    });
    
    // Save updated history
    scriptProps.setProperty('sessionHistory', JSON.stringify(historyArray));
    
    // Refresh menu system
    if (AdministrativeModule && typeof AdministrativeModule.createMenu === 'function') {
      AdministrativeModule.createMenu();
    }
    
    return true;
  } catch (error) {
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'finalizeTransition', 
        'Error finalizing transition. Please try again or contact support.');
    } else {
      Logger.log(`Error finalizing transition: ${error.message}`);
      SpreadsheetApp.getUi().alert(
        'Error',
        `Failed to finalize transition: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
    // Update status
    updateTransitionStatus(TRANSITION_STEPS.FINALIZE, 'Failed');
    
    return false;
  }
}

// Global variable export
const SessionTransitionModule = {
  startSessionTransition,
  resumeSessionTransition
};