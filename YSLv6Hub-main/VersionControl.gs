/**
 * YSL Hub v2 Version Control Module
 * 
 * This module manages system versions, updates, and diagnostics.
 * It provides a centralized way to track versions and apply updates.
 * 
 * @author PenBay YMCA
 * @version 1.0
 * @date 2025-04-27
 */

// Version constants
const VERSION = {
  CURRENT: '2.0.0',
  MIN_COMPATIBLE: '1.0.0',
  RELEASE_DATE: '2025-04-27'
};

// Script properties keys
const VERSION_PROPS = {
  CURRENT_VERSION: 'app_version',
  LAST_UPDATE: 'last_update_date',
  UPDATE_LOG: 'update_log'
};

/**
 * Initialize version control system
 * Sets initial version information if not already present
 */
function initializeVersionControl() {
  const scriptProps = PropertiesService.getScriptProperties();
  
  // If version not set, initialize it
  if (!scriptProps.getProperty(VERSION_PROPS.CURRENT_VERSION)) {
    scriptProps.setProperty(VERSION_PROPS.CURRENT_VERSION, VERSION.CURRENT);
    scriptProps.setProperty(VERSION_PROPS.LAST_UPDATE, new Date().toISOString());
    scriptProps.setProperty(VERSION_PROPS.UPDATE_LOG, JSON.stringify([{
      date: new Date().toISOString(),
      version: VERSION.CURRENT,
      notes: 'Initial system installation'
    }]));
    
    logMessage('Version control initialized', 'INFO');
  }
}

/**
 * Get the current system version
 * @return {string} Current version number
 */
function getCurrentVersion() {
  const scriptProps = PropertiesService.getScriptProperties();
  return scriptProps.getProperty(VERSION_PROPS.CURRENT_VERSION) || VERSION.CURRENT;
}

/**
 * Get version information including current version and update history
 * @return {Object} Version information
 */
function getVersionInfo() {
  const scriptProps = PropertiesService.getScriptProperties();
  const currentVersion = scriptProps.getProperty(VERSION_PROPS.CURRENT_VERSION) || VERSION.CURRENT;
  const lastUpdate = scriptProps.getProperty(VERSION_PROPS.LAST_UPDATE) || new Date().toISOString();
  
  let updateLog = [];
  try {
    const updateLogJson = scriptProps.getProperty(VERSION_PROPS.UPDATE_LOG);
    if (updateLogJson) {
      updateLog = JSON.parse(updateLogJson);
    }
  } catch (error) {
    logMessage(`Error parsing update log: ${error.message}`, 'ERROR');
    updateLog = [];
  }
  
  return {
    currentVersion: currentVersion,
    releaseDate: VERSION.RELEASE_DATE,
    lastUpdate: lastUpdate,
    minCompatibleVersion: VERSION.MIN_COMPATIBLE,
    updateHistory: updateLog
  };
}

/**
 * Record a new version update
 * @param {string} version - New version number
 * @param {string} notes - Update notes
 */
function recordUpdate(version, notes) {
  const scriptProps = PropertiesService.getScriptProperties();
  
  // Update current version
  scriptProps.setProperty(VERSION_PROPS.CURRENT_VERSION, version);
  scriptProps.setProperty(VERSION_PROPS.LAST_UPDATE, new Date().toISOString());
  
  // Add to update log
  let updateLog = [];
  try {
    const updateLogJson = scriptProps.getProperty(VERSION_PROPS.UPDATE_LOG);
    if (updateLogJson) {
      updateLog = JSON.parse(updateLogJson);
    }
  } catch (error) {
    logMessage(`Error parsing update log: ${error.message}`, 'ERROR');
    updateLog = [];
  }
  
  // Add new update record
  updateLog.push({
    date: new Date().toISOString(),
    version: version,
    notes: notes
  });
  
  // Save update log
  scriptProps.setProperty(VERSION_PROPS.UPDATE_LOG, JSON.stringify(updateLog));
  
  logMessage(`System updated to version ${version}: ${notes}`, 'INFO');
}

/**
 * Display version information to the user
 */
function showVersionInfo() {
  const versionInfo = getVersionInfo();
  const ui = SpreadsheetApp.getUi();
  
  let updateHistoryText = '';
  if (versionInfo.updateHistory && versionInfo.updateHistory.length > 0) {
    // Sort updates from newest to oldest
    const sortedUpdates = versionInfo.updateHistory.sort((a, b) => 
      new Date(b.date).getTime() - new Date(a.date).getTime()
    );
    
    // Format the last 5 updates (or all if less than 5)
    const recentUpdates = sortedUpdates.slice(0, 5);
    updateHistoryText = '\n\nRecent Updates:\n';
    
    recentUpdates.forEach(update => {
      const updateDate = new Date(update.date);
      const formattedDate = Utilities.formatDate(updateDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
      updateHistoryText += `- ${formattedDate}: Version ${update.version} - ${update.notes}\n`;
    });
  }
  
  const lastUpdateDate = new Date(versionInfo.lastUpdate);
  const formattedLastUpdate = Utilities.formatDate(lastUpdateDate, Session.getScriptTimeZone(), 'MMM dd, yyyy');
  
  ui.alert(
    'YSL Hub v2 - Version Information',
    `Current Version: ${versionInfo.currentVersion}\n` +
    `Release Date: ${versionInfo.releaseDate}\n` +
    `Last Updated: ${formattedLastUpdate}` +
    updateHistoryText,
    ui.ButtonSet.OK
  );
}

/**
 * Check if the system needs an update (for future use)
 * @return {boolean} True if system needs update
 */
function checkForUpdates() {
  // This is a placeholder for future update checking functionality
  // In a production environment, this could check against a remote source
  
  return false;
}

/**
 * Run system diagnostics and report status
 * @return {Object} System status information
 */
function runDiagnostics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const diagnostics = {
    version: getCurrentVersion(),
    timestamp: new Date().toISOString(),
    status: 'OK',
    issues: [],
    checks: []
  };
  
  // Check for required sheets
  const requiredSheets = ['Assumptions', 'Daxko', 'Classes'];
  for (const sheetName of requiredSheets) {
    const sheet = ss.getSheetByName(sheetName);
    const checkResult = {
      name: `Required Sheet: ${sheetName}`,
      status: sheet ? 'PASS' : 'FAIL',
      details: sheet ? `Sheet "${sheetName}" exists` : `Sheet "${sheetName}" not found`
    };
    
    diagnostics.checks.push(checkResult);
    if (!sheet) {
      diagnostics.status = 'WARNING';
      diagnostics.issues.push(`Required sheet "${sheetName}" not found`);
    }
  }
  
  // Check system configuration
  const config = PropertiesService.getScriptProperties().getProperties();
  const requiredProps = ['sessionName', 'rosterFolderUrl', 'swimmerRecordsUrl'];
  
  for (const prop of requiredProps) {
    const checkResult = {
      name: `Configuration: ${prop}`,
      status: config[prop] ? 'PASS' : 'FAIL',
      details: config[prop] ? `Property "${prop}" is set` : `Property "${prop}" not set`
    };
    
    diagnostics.checks.push(checkResult);
    if (!config[prop]) {
      diagnostics.status = 'WARNING';
      diagnostics.issues.push(`Required configuration "${prop}" not set`);
    }
  }
  
  // Log diagnostics result
  logMessage(`System diagnostics run: ${diagnostics.status}`, 'INFO');
  if (diagnostics.issues.length > 0) {
    for (const issue of diagnostics.issues) {
      logMessage(`Diagnostic issue: ${issue}`, 'WARNING');
    }
  }
  
  return diagnostics;
}

/**
 * Display diagnostics results to the user
 */
function showDiagnostics() {
  const ui = SpreadsheetApp.getUi();
  const diagnostics = runDiagnostics();
  
  let checksText = '';
  for (const check of diagnostics.checks) {
    checksText += `${check.name}: ${check.status}\n  ${check.details}\n`;
  }
  
  let issuesText = '';
  if (diagnostics.issues.length > 0) {
    issuesText = '\nIssues Found:\n';
    for (const issue of diagnostics.issues) {
      issuesText += `- ${issue}\n`;
    }
  }
  
  ui.alert(
    'YSL Hub System Diagnostics',
    `Status: ${diagnostics.status}\n` +
    `Version: ${diagnostics.version}\n\n` +
    'Checks:\n' + checksText +
    issuesText +
    '\nIf issues were found, please contact the system administrator.',
    ui.ButtonSet.OK
  );
}

/**
 * Log a message to the system log
 * This function is a bridge to the ErrorHandling module's logMessage function
 * 
 * @param {string} message - Message to log
 * @param {string} level - Log level (INFO, WARNING, ERROR)
 */
function logMessage(message, level = 'INFO') {
  // If ErrorHandling module exists, use its logMessage function
  if (typeof ErrorHandling !== 'undefined' && typeof ErrorHandling.logMessage === 'function') {
    ErrorHandling.logMessage(message, level);
  } else {
    // Fallback to basic logging
    Logger.log(`[${level}] ${message}`);
  }
}

// Make functions available to other modules
const VersionControl = {
  initializeVersionControl: initializeVersionControl,
  getCurrentVersion: getCurrentVersion,
  getVersionInfo: getVersionInfo,
  recordUpdate: recordUpdate,
  showVersionInfo: showVersionInfo,
  checkForUpdates: checkForUpdates,
  runDiagnostics: runDiagnostics,
  showDiagnostics: showDiagnostics
};