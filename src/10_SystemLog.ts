/**
 * YSLv6Hub SystemLog Module
 * 
 * This module provides enhanced logging functionality with structured log entries,
 * severity levels, and searchable logs.
 * 
 * @author Sean R. Sullivan
 * @version 1.0.0
 * @date 2025-05-18
 */

import { ErrorSeverity } from './01_Core';

/**
 * Log entry interface
 */
export interface LogEntry {
  timestamp: string;
  severity: ErrorSeverity;
  module: string;
  function: string;
  message: string;
  details?: any;
  user?: string;
}

/**
 * Search criteria interface for log searching
 */
export interface LogSearchCriteria {
  module?: string;
  severity?: ErrorSeverity;
  dateFrom?: Date;
  dateTo?: Date;
  messageContains?: string;
}

/**
 * SystemLog for enhanced logging functionality
 */
export const SystemLog = {
  /**
   * Logs an entry to the SystemLog sheet
   */
  log(entry: LogEntry): void {
    try {
      // Ensure timestamp exists
      if (!entry.timestamp) {
        entry.timestamp = new Date().toISOString();
      }
      
      // Add user if not present
      if (!entry.user) {
        entry.user = Session.getActiveUser().getEmail();
      }
      
      // Get SystemLog sheet
      const sheet = this.getOrCreateSystemLogSheet();
      
      // Prepare log row
      const logRow = [
        entry.timestamp,
        entry.severity,
        entry.module,
        entry.function,
        entry.message,
        JSON.stringify(entry.details || {}),
        entry.user
      ];
      
      // Add row to sheet
      sheet.appendRow(logRow);
      
      // If log gets too large, trim it
      this.trimLogIfNeeded(sheet);
      
    } catch (error) {
      console.error('Error logging to SystemLog:', error);
    }
  },
  
  /**
   * Gets or creates the SystemLog sheet
   */
  getOrCreateSystemLogSheet(): GoogleAppsScript.Spreadsheet.Sheet {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('SystemLog');
    
    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('SystemLog');
      
      // Set up header row
      sheet.appendRow([
        'Timestamp',
        'Severity',
        'Module',
        'Function',
        'Message',
        'Details',
        'User'
      ]);
      
      // Format header row
      sheet.getRange(1, 1, 1, 7)
        .setBackground('#f3f3f3')
        .setFontWeight('bold');
      
      // Freeze header row
      sheet.setFrozenRows(1);
      
      // Adjust column widths for readability
      sheet.setColumnWidth(1, 150); // Timestamp
      sheet.setColumnWidth(2, 80);  // Severity
      sheet.setColumnWidth(3, 120); // Module
      sheet.setColumnWidth(4, 120); // Function
      sheet.setColumnWidth(5, 300); // Message
      sheet.setColumnWidth(6, 200); // Details
      sheet.setColumnWidth(7, 150); // User
    }
    
    return sheet;
  },
  
  /**
   * Trims the log if it gets too large
   */
  trimLogIfNeeded(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    const MAX_LOG_ROWS = 10000;
    const currentRows = sheet.getLastRow();
    
    if (currentRows > MAX_LOG_ROWS) {
      // Keep header row and the most recent entries
      const rowsToDelete = currentRows - MAX_LOG_ROWS;
      sheet.deleteRows(2, rowsToDelete);
    }
  },
  
  /**
   * Logs an info message
   */
  info(module: string, functionName: string, message: string, details?: any): void {
    this.log({
      timestamp: new Date().toISOString(),
      severity: ErrorSeverity.INFO,
      module,
      function: functionName,
      message,
      details
    });
  },
  
  /**
   * Logs a warning message
   */
  warning(module: string, functionName: string, message: string, details?: any): void {
    this.log({
      timestamp: new Date().toISOString(),
      severity: ErrorSeverity.WARNING,
      module,
      function: functionName,
      message,
      details
    });
  },
  
  /**
   * Logs an error message
   */
  error(module: string, functionName: string, message: string, details?: any): void {
    this.log({
      timestamp: new Date().toISOString(),
      severity: ErrorSeverity.ERROR,
      module,
      function: functionName,
      message,
      details
    });
  },
  
  /**
   * Logs a critical error message
   */
  critical(module: string, functionName: string, message: string, details?: any): void {
    this.log({
      timestamp: new Date().toISOString(),
      severity: ErrorSeverity.CRITICAL,
      module,
      function: functionName,
      message,
      details
    });
  },
  
  /**
   * Searches the log for entries matching criteria
   */
  search(_criteria: LogSearchCriteria): LogEntry[] {
    // Implementation will search logs based on criteria
    // Currently a placeholder that will be implemented later
    return [];
  },
  
  /**
   * Clears the log sheet
   */
  clearLog(): void {
    const sheet = this.getOrCreateSystemLogSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
  }
};

/**
 * Shows the system log viewer
 */
function showSystemLog(): void {
  // Implementation will show a log viewer UI
  // Currently a placeholder
  SpreadsheetApp.getUi().alert('System Log Viewer', 'System log viewer is not yet implemented.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// Make functions available globally for menu actions
// @ts-ignore
global.showSystemLog = showSystemLog;