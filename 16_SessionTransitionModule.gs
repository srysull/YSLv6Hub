/**
 * YSL Hub Session Transition Module
 * 
 * This module handles the transition between sessions, providing a guided
 * workflow for archiving current session data and setting up a new session.
 * 
 * @author Sean R. Sullivan
 * @version 1.0
 * @date 2025-05-14
 */

// Configuration constants
const SESSION_TRANSITION_CONFIG = {
  PROPS: {
    TRANSITION_STATE: 'transition_state',
    TRANSITION_STEP: 'transition_step',
    PREVIOUS_SESSION: 'previous_session',
    NEW_SESSION: 'new_session',
    BACKUP_CONFIG: 'config_backup'
  },
  STATES: {
    IDLE: 'idle',
    IN_PROGRESS: 'in_progress',
    COMPLETED: 'completed',
    FAILED: 'failed'
  },
  STEPS: {
    NOT_STARTED: 0,
    BACKUP_CREATED: 1,
    DATA_ARCHIVED: 2,
    NEW_SESSION_INITIALIZED: 3,
    COMPLETED: 4
  }
};

/**
 * Starts the session transition process with a guided wizard
 * @return {boolean} Success status
 */
function startSessionTransition() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Check if a transition is already in progress
    const transitionState = PropertiesService.getDocumentProperties()
      .getProperty(SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STATE);
    
    if (transitionState === SESSION_TRANSITION_CONFIG.STATES.IN_PROGRESS) {
      // Offer to resume the transition
      const response = ui.alert(
        'Transition In Progress',
        'A session transition is already in progress. Would you like to resume it?',
        ui.ButtonSet.YES_NO
      );
      
      if (response === ui.Button.YES) {
        return resumeSessionTransition();
      } else {
        return false;
      }
    }
    
    // Get the current session name
    const currentSession = GlobalFunctions.safeGetProperty(CONFIG.SESSION_NAME, '');
    
    if (!currentSession) {
      ui.alert(
        'Session Not Found',
        'Could not determine the current session name. Please ensure the system is properly initialized.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Display the transition checklist
    const checklistResponse = showTransitionChecklist(currentSession);
    
    if (checklistResponse === ui.Button.CANCEL) {
      return false;
    }
    
    // Get new session information
    const sessionResponse = ui.prompt(
      'Start New Session - Step 1 of 3',
      'Please enter a name for the new session (e.g., "Summer 2025"):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (sessionResponse.getSelectedButton() === ui.Button.CANCEL) {
      return false;
    }
    
    const newSessionName = sessionResponse.getResponseText().trim();
    
    if (!newSessionName) {
      ui.alert(
        'Invalid Session Name',
        'Please provide a valid session name to continue.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Get roster folder URL
    const rosterResponse = ui.prompt(
      'Start New Session - Step 2 of 3',
      'Please enter the URL for the roster folder for the new session:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (rosterResponse.getSelectedButton() === ui.Button.CANCEL) {
      return false;
    }
    
    const rosterFolderUrl = rosterResponse.getResponseText().trim();
    
    if (!rosterFolderUrl) {
      ui.alert(
        'Invalid Roster Folder URL',
        'Please provide a valid URL for the roster folder to continue.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Ask which data to carry forward (most recent settings are default)
    const dataChoicesResponse = ui.alert(
      'Start New Session - Step 3 of 3',
      'Would you like to carry forward system settings from the current session?\n\n' +
      'This includes report templates, swimmer records, and handbook URLs.',
      ui.ButtonSet.YES_NO
    );
    
    const carryForwardSettings = (dataChoicesResponse === ui.Button.YES);
    
    // Confirm the transition
    const confirmResponse = ui.alert(
      'Confirm Session Transition',
      `You are about to transition from "${currentSession}" to "${newSessionName}".\n\n` +
      'This will:\n' +
      '1. Archive all current session data\n' +
      '2. Initialize the system for the new session\n' +
      (carryForwardSettings ? '3. Carry forward system settings from the current session\n\n' : '\n') +
      'Are you sure you want to continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (confirmResponse === ui.Button.NO) {
      return false;
    }
    
    // Start the transition process
    return executeSessionTransition(currentSession, newSessionName, rosterFolderUrl, carryForwardSettings);
  } catch (error) {
    console.error('Error starting session transition: ' + error.message);
    
    // Log detailed error for troubleshooting
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Session transition failed to start: ' + error.message, 'ERROR', 'startSessionTransition');
    }
    
    // Show a user-friendly error
    SpreadsheetApp.getUi().alert(
      'Transition Error',
      'Could not start the session transition. Please try again or contact support.\n\n' + 
      'Error details: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false;
  }
}

/**
 * Shows a transition checklist to ensure all tasks are completed
 * @param {string} currentSession - The current session name
 * @return {Button} The user's response (OK or CANCEL)
 */
function showTransitionChecklist(currentSession) {
  const ui = SpreadsheetApp.getUi();
  
  return ui.alert(
    'End of Session Checklist',
    `Before transitioning from "${currentSession}", please confirm you've completed these tasks:\n\n` +
    '✓ All student assessments have been completed in instructor sheets\n' +
    '✓ End-session reports have been sent to parents/guardians\n' +
    '✓ You have access to the roster folder for the new session\n' +
    '✓ You have backed up any important data from the current session\n\n' +
    'Have you completed all these tasks?',
    ui.ButtonSet.OK_CANCEL
  );
}

/**
 * Executes the session transition process
 * @param {string} currentSession - The current session name
 * @param {string} newSession - The new session name
 * @param {string} rosterFolderUrl - The URL for the roster folder
 * @param {boolean} carryForwardSettings - Whether to carry forward settings
 * @return {boolean} Success status
 */
function executeSessionTransition(currentSession, newSession, rosterFolderUrl, carryForwardSettings) {
  try {
    // Initialize transition state
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STATE, 
      SESSION_TRANSITION_CONFIG.STATES.IN_PROGRESS
    );
    
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
      SESSION_TRANSITION_CONFIG.STEPS.NOT_STARTED.toString()
    );
    
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.PREVIOUS_SESSION, 
      currentSession
    );
    
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.NEW_SESSION, 
      newSession
    );
    
    // Step 1: Backup current configuration
    const configBackup = backupCurrentConfiguration();
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.BACKUP_CONFIG, 
      JSON.stringify(configBackup)
    );
    
    // Update transition step
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
      SESSION_TRANSITION_CONFIG.STEPS.BACKUP_CREATED.toString()
    );
    
    // Step 2: Archive current session data
    const archiveSuccess = archiveCurrentSessionData(currentSession);
    
    if (!archiveSuccess) {
      throw new Error('Failed to archive current session data');
    }
    
    // Update transition step
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
      SESSION_TRANSITION_CONFIG.STEPS.DATA_ARCHIVED.toString()
    );
    
    // Step 3: Initialize new session
    const initializeSuccess = initializeNewSession(newSession, rosterFolderUrl, carryForwardSettings, configBackup);
    
    if (!initializeSuccess) {
      throw new Error('Failed to initialize new session');
    }
    
    // Update transition step
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
      SESSION_TRANSITION_CONFIG.STEPS.NEW_SESSION_INITIALIZED.toString()
    );
    
    // Step 4: Complete the transition
    const cleanupSuccess = cleanupTransition();
    
    if (!cleanupSuccess) {
      throw new Error('Failed to complete transition cleanup');
    }
    
    // Mark transition as completed
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STATE, 
      SESSION_TRANSITION_CONFIG.STATES.COMPLETED
    );
    
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
      SESSION_TRANSITION_CONFIG.STEPS.COMPLETED.toString()
    );
    
    // Show success message
    SpreadsheetApp.getUi().alert(
      'Session Transition Complete',
      `Successfully transitioned from "${currentSession}" to "${newSession}".\n\n` +
      'The previous session data has been archived and can be accessed in the History sheet.\n\n' +
      'The system is now initialized for the new session.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    console.error('Error executing session transition: ' + error.message);
    
    // Mark transition as failed
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STATE, 
      SESSION_TRANSITION_CONFIG.STATES.FAILED
    );
    
    // Log detailed error for troubleshooting
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Session transition failed: ' + error.message, 'ERROR', 'executeSessionTransition');
    }
    
    // Show a user-friendly error
    SpreadsheetApp.getUi().alert(
      'Transition Error',
      'The session transition encountered an error. You can resume the transition later.\n\n' + 
      'Error details: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false;
  }
}

/**
 * Resumes a previously started session transition
 * @return {boolean} Success status
 */
function resumeSessionTransition() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get transition state information
    const docProps = PropertiesService.getDocumentProperties();
    const transitionStep = parseInt(docProps.getProperty(SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP) || '0');
    const previousSession = docProps.getProperty(SESSION_TRANSITION_CONFIG.PROPS.PREVIOUS_SESSION);
    const newSession = docProps.getProperty(SESSION_TRANSITION_CONFIG.PROPS.NEW_SESSION);
    
    if (!previousSession || !newSession) {
      ui.alert(
        'Transition Data Missing',
        'Could not find the previous or new session information. Please start a new transition.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Show the current state
    const statusMessage = getTransitionStatusMessage(transitionStep, previousSession, newSession);
    
    const resumeResponse = ui.alert(
      'Resume Session Transition',
      statusMessage + '\n\nWould you like to resume this transition?',
      ui.ButtonSet.YES_NO
    );
    
    if (resumeResponse === ui.Button.NO) {
      return false;
    }
    
    // Get the backup configuration
    let configBackup = {};
    try {
      const backupJson = docProps.getProperty(SESSION_TRANSITION_CONFIG.PROPS.BACKUP_CONFIG);
      if (backupJson) {
        configBackup = JSON.parse(backupJson);
      }
    } catch (e) {
      console.warn('Could not parse config backup: ' + e.message);
      // Continue with empty config
    }
    
    // Resume from the appropriate step
    let success = false;
    
    if (transitionStep < SESSION_TRANSITION_CONFIG.STEPS.BACKUP_CREATED) {
      // Start from backup step
      configBackup = backupCurrentConfiguration();
      docProps.setProperty(
        SESSION_TRANSITION_CONFIG.PROPS.BACKUP_CONFIG, 
        JSON.stringify(configBackup)
      );
      
      docProps.setProperty(
        SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
        SESSION_TRANSITION_CONFIG.STEPS.BACKUP_CREATED.toString()
      );
    }
    
    if (transitionStep < SESSION_TRANSITION_CONFIG.STEPS.DATA_ARCHIVED) {
      // Archive the current session
      success = archiveCurrentSessionData(previousSession);
      
      if (!success) {
        throw new Error('Failed to archive current session data');
      }
      
      docProps.setProperty(
        SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
        SESSION_TRANSITION_CONFIG.STEPS.DATA_ARCHIVED.toString()
      );
    }
    
    if (transitionStep < SESSION_TRANSITION_CONFIG.STEPS.NEW_SESSION_INITIALIZED) {
      // Ask for roster URL if needed
      let rosterFolderUrl = '';
      const rosterResponse = ui.prompt(
        'Resume Transition',
        'Please enter the URL for the roster folder for the new session:',
        ui.ButtonSet.OK_CANCEL
      );
      
      if (rosterResponse.getSelectedButton() === ui.Button.CANCEL) {
        return false;
      }
      
      rosterFolderUrl = rosterResponse.getResponseText().trim();
      
      if (!rosterFolderUrl) {
        ui.alert(
          'Invalid Roster Folder URL',
          'Please provide a valid URL for the roster folder to continue.',
          ui.ButtonSet.OK
        );
        return false;
      }
      
      // Ask about carrying forward settings
      const settingsResponse = ui.alert(
        'Resume Transition',
        'Would you like to carry forward system settings from the previous session?',
        ui.ButtonSet.YES_NO
      );
      
      const carryForwardSettings = (settingsResponse === ui.Button.YES);
      
      // Initialize the new session
      success = initializeNewSession(newSession, rosterFolderUrl, carryForwardSettings, configBackup);
      
      if (!success) {
        throw new Error('Failed to initialize new session');
      }
      
      docProps.setProperty(
        SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
        SESSION_TRANSITION_CONFIG.STEPS.NEW_SESSION_INITIALIZED.toString()
      );
    }
    
    // Complete the transition
    success = cleanupTransition();
    
    if (!success) {
      throw new Error('Failed to complete transition cleanup');
    }
    
    // Mark transition as completed
    docProps.setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STATE, 
      SESSION_TRANSITION_CONFIG.STATES.COMPLETED
    );
    
    docProps.setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STEP, 
      SESSION_TRANSITION_CONFIG.STEPS.COMPLETED.toString()
    );
    
    // Show success message
    ui.alert(
      'Session Transition Complete',
      `Successfully transitioned from "${previousSession}" to "${newSession}".\n\n` +
      'The previous session data has been archived and can be accessed in the History sheet.\n\n' +
      'The system is now initialized for the new session.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    console.error('Error resuming session transition: ' + error.message);
    
    // Mark transition as failed
    PropertiesService.getDocumentProperties().setProperty(
      SESSION_TRANSITION_CONFIG.PROPS.TRANSITION_STATE, 
      SESSION_TRANSITION_CONFIG.STATES.FAILED
    );
    
    // Log detailed error for troubleshooting
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Session transition resume failed: ' + error.message, 'ERROR', 'resumeSessionTransition');
    }
    
    // Show a user-friendly error
    SpreadsheetApp.getUi().alert(
      'Transition Resume Error',
      'Could not resume the session transition. Please try again or contact support.\n\n' + 
      'Error details: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return false;
  }
}

/**
 * Gets a user-friendly status message for the current transition state
 * @param {number} step - The current transition step
 * @param {string} previousSession - The previous session name
 * @param {string} newSession - The new session name
 * @return {string} A status message
 */
function getTransitionStatusMessage(step, previousSession, newSession) {
  let message = `Transitioning from "${previousSession}" to "${newSession}"\n\n`;
  
  message += 'Current progress:\n';
  
  if (step >= SESSION_TRANSITION_CONFIG.STEPS.BACKUP_CREATED) {
    message += '✓ Configuration backup created\n';
  } else {
    message += '○ Configuration backup pending\n';
  }
  
  if (step >= SESSION_TRANSITION_CONFIG.STEPS.DATA_ARCHIVED) {
    message += '✓ Current session data archived\n';
  } else {
    message += '○ Session data archiving pending\n';
  }
  
  if (step >= SESSION_TRANSITION_CONFIG.STEPS.NEW_SESSION_INITIALIZED) {
    message += '✓ New session initialized\n';
  } else {
    message += '○ New session initialization pending\n';
  }
  
  if (step >= SESSION_TRANSITION_CONFIG.STEPS.COMPLETED) {
    message += '✓ Transition cleanup completed\n';
  } else {
    message += '○ Transition cleanup pending\n';
  }
  
  return message;
}

/**
 * Backs up the current system configuration
 * @return {Object} The configuration backup
 */
function backupCurrentConfiguration() {
  try {
    // Get all script properties
    const allProps = PropertiesService.getScriptProperties().getProperties();
    
    // Create a backup object
    const backup = {
      timestamp: new Date().toISOString(),
      properties: allProps
    };
    
    console.log('Configuration backup created');
    return backup;
  } catch (error) {
    console.error('Error backing up configuration: ' + error.message);
    throw error;
  }
}

/**
 * Archives the current session data
 * @param {string} sessionName - The current session name
 * @return {boolean} Success status
 */
function archiveCurrentSessionData(sessionName) {
  try {
    // Use the HistoryModule to archive the current session
    if (typeof HistoryModule !== 'undefined' && typeof HistoryModule.archiveCurrentSession === 'function') {
      const success = HistoryModule.archiveCurrentSession(sessionName);
      
      if (!success) {
        throw new Error('History module reported failure during archiving');
      }
    } else {
      throw new Error('History module not available for archiving');
    }
    
    console.log('Current session data archived successfully');
    return true;
  } catch (error) {
    console.error('Error archiving session data: ' + error.message);
    
    // Log detailed error for troubleshooting
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Session archiving failed: ' + error.message, 'ERROR', 'archiveCurrentSessionData');
    }
    
    // Warn the user but don't stop the transition
    SpreadsheetApp.getUi().alert(
      'Archiving Warning',
      'Some data may not have been fully archived, but the transition will continue.\n\n' + 
      'Error details: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    // Return true to allow transition to continue
    return true;
  }
}

/**
 * Initializes the system for a new session
 * @param {string} newSession - The new session name
 * @param {string} rosterFolderUrl - The URL for the roster folder
 * @param {boolean} carryForwardSettings - Whether to carry forward settings
 * @param {Object} configBackup - The configuration backup
 * @return {boolean} Success status
 */
function initializeNewSession(newSession, rosterFolderUrl, carryForwardSettings, configBackup) {
  try {
    // Clear current session properties
    GlobalFunctions.safeSetProperty(CONFIG.SESSION_NAME, newSession);
    GlobalFunctions.safeSetProperty(CONFIG.ROSTER_FOLDER_URL, rosterFolderUrl);
    
    // Carry forward other settings if requested
    if (carryForwardSettings && configBackup && configBackup.properties) {
      const oldProps = configBackup.properties;
      
      // Carry forward these specific properties if they exist
      const propsToCarry = [
        CONFIG.REPORT_TEMPLATE_URL,
        CONFIG.SWIMMER_RECORDS_URL,
        CONFIG.PARENT_HANDBOOK_URL,
        CONFIG.SESSION_PROGRAMS_URL
      ];
      
      for (const prop of propsToCarry) {
        if (oldProps[prop]) {
          GlobalFunctions.safeSetProperty(prop, oldProps[prop]);
        }
      }
    }
    
    // Initialize data structures for the new session
    if (typeof DataIntegrationModule !== 'undefined' && 
        typeof DataIntegrationModule.initializeDataStructures === 'function') {
      
      // Build config object for initialization
      const config = {
        sessionName: newSession,
        rosterFolderUrl: rosterFolderUrl,
        reportTemplateUrl: GlobalFunctions.safeGetProperty(CONFIG.REPORT_TEMPLATE_URL),
        swimmerRecordsUrl: GlobalFunctions.safeGetProperty(CONFIG.SWIMMER_RECORDS_URL),
        parentHandbookUrl: GlobalFunctions.safeGetProperty(CONFIG.PARENT_HANDBOOK_URL),
        sessionProgramsUrl: GlobalFunctions.safeGetProperty(CONFIG.SESSION_PROGRAMS_URL)
      };
      
      // Initialize data structures
      DataIntegrationModule.initializeDataStructures(config);
    } else {
      throw new Error('Data Integration module not available for initialization');
    }
    
    // Mark system as initialized
    GlobalFunctions.safeSetProperty(CONFIG.INITIALIZED, 'true');
    
    console.log('New session initialized successfully');
    return true;
  } catch (error) {
    console.error('Error initializing new session: ' + error.message);
    
    // Log detailed error for troubleshooting
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('New session initialization failed: ' + error.message, 'ERROR', 'initializeNewSession');
    }
    
    throw error;
  }
}

/**
 * Cleans up after transition is complete
 * @return {boolean} Success status
 */
function cleanupTransition() {
  try {
    // Refresh the menu
    AdministrativeModule.createMenu();
    
    // Reset transition properties (but keep a record of the last transition)
    const docProps = PropertiesService.getDocumentProperties();
    docProps.setProperty('LAST_TRANSITION_DATE', new Date().toISOString());
    docProps.setProperty('LAST_TRANSITION_FROM', docProps.getProperty(SESSION_TRANSITION_CONFIG.PROPS.PREVIOUS_SESSION));
    docProps.setProperty('LAST_TRANSITION_TO', docProps.getProperty(SESSION_TRANSITION_CONFIG.PROPS.NEW_SESSION));
    
    // Clear active transition properties
    docProps.deleteProperty(SESSION_TRANSITION_CONFIG.PROPS.BACKUP_CONFIG);
    
    console.log('Transition cleanup completed successfully');
    return true;
  } catch (error) {
    console.error('Error in transition cleanup: ' + error.message);
    
    // Log but don't fail the operation
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Transition cleanup issue: ' + error.message, 'WARNING', 'cleanupTransition');
    }
    
    // Return true even if cleanup has issues
    return true;
  }
}

// Make functions available to other modules
const SessionTransitionModule = {
  startSessionTransition: startSessionTransition,
  resumeSessionTransition: resumeSessionTransition
};