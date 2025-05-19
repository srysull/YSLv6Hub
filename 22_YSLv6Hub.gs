/**
 * YSL Hub v2 YSLv6Hub Dashboard Module
 * 
 * This module handles the setup and management of the YSLv6Hub dashboard.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-14
 */

/**
 * Sets up the YSLv6Hub dashboard with system status tracking
 * @param {Sheet} sheet - The YSLv6Hub sheet to set up
 */
function setupYSLv6Hub(sheet) {
  try {
    // Clear existing content
    sheet.clear();
    
    // Create dashboard header
    sheet.getRange('A1:C1').merge();
    sheet.getRange('A1').setValue('YSLv6Hub Dashboard')
      .setFontWeight('bold')
      .setFontSize(16)
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    
    // Section 1: Session Information
    sheet.getRange('A3:C3').merge();
    sheet.getRange('A3').setValue('Session Information')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#e9f2fe');
    
    // Session Name Input
    sheet.getRange('A4').setValue('Session Name (Required):')
      .setFontWeight('bold');
    sheet.getRange('B4').setValue('')
      .setBackground('#f0f0f0')
      .setNote('Enter the name of the current session (e.g., "Summer 2025")');
    
    // Section 2: System Initialization Status
    sheet.getRange('A6:C6').merge();
    sheet.getRange('A6').setValue('System Initialization Status')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#e9f2fe');
    
    // Status table headers
    sheet.getRange('A7:C7').setValues([['Action', 'Status', 'Notes']])
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    
    // Status table rows
    const statusRows = [
      ['Import Registration Data', 'Pending', 'Use "Import Registration Data" from menu'],
      ['Map Registration Fields', 'Pending', 'Complete field mapping below'],
      ['Generate SwimmerSkills', 'Pending', 'Use menu after field mapping'],
      ['Generate SwimmerLog', 'Pending', 'Use menu after field mapping'],
      ['Generate Groups Tracker', 'Pending', 'Use menu after field mapping'],
      ['Generate Communications Hub', 'Pending', 'Use menu after field mapping'],
      ['Generate Privates Tracker', 'Pending', 'Use menu after field mapping'],
      ['Sync Groups Tracker Info', 'Pending', 'Use menu after creating trackers'],
      ['Sync Privates Tracker Info', 'Pending', 'Use menu after creating trackers'],
      ['Session Transition', 'Pending', 'Use menu at end of session']
    ];
    
    sheet.getRange(8, 1, statusRows.length, 3).setValues(statusRows);
    
    // Format status cells
    for (let i = 0; i < statusRows.length; i++) {
      sheet.getRange(8 + i, 2).setBackground('#ffe0b2'); // Orange for pending
    }
    
    // Section 3: Field Mappings - This will be populated when RegInfo is imported
    sheet.getRange('A19:C19').merge();
    sheet.getRange('A19').setValue('Field Mappings (Required)')
      .setFontWeight('bold')
      .setFontSize(12)
      .setBackground('#e9f2fe');
    
    // Field mapping instructions
    sheet.getRange('A20:C20').merge();
    sheet.getRange('A20').setValue('First import registration data to enable field mapping')
      .setFontStyle('italic');
    
    // Set column widths
    sheet.setColumnWidth(1, 200); // Field Name
    sheet.setColumnWidth(2, 200); // Value/RegInfo Column
    sheet.setColumnWidth(3, 300); // Description
    
    // Add note about workflow
    sheet.getRange('A22:C22').merge();
    sheet.getRange('A22').setValue('System workflow: Import data → Map fields → Generate trackers → Sync data')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setBackground('#e6f2ff');
    
  } catch (error) {
    Logger.log('Error setting up YSLv6Hub dashboard: ' + error);
  }
}

/**
 * Updates the initialization status for a specific action
 * @param {string} action - The action to update
 * @param {string} status - The new status ('Complete', 'Pending', or 'In Progress')
 */
function updateInitializationStatus(action, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('YSLv6Hub');
    
    if (!dashboardSheet) {
      Logger.log('YSLv6Hub sheet not found');
      return;
    }
    
    // Find the row with the specified action
    const data = dashboardSheet.getDataRange().getValues();
    let actionRow = -1;
    
    for (let i = 7; i < data.length; i++) {
      if (data[i][0] === action) {
        actionRow = i + 1; // +1 because rows are 1-indexed
        break;
      }
    }
    
    if (actionRow === -1) {
      return; // Action not found
    }
    
    // Update the status
    dashboardSheet.getRange(actionRow, 2).setValue(status);
    
    // Set background color based on status
    if (status === 'Complete') {
      dashboardSheet.getRange(actionRow, 2).setBackground('#c8e6c9'); // Green
    } else if (status === 'In Progress') {
      dashboardSheet.getRange(actionRow, 2).setBackground('#bbdefb'); // Blue
    } else {
      dashboardSheet.getRange(actionRow, 2).setBackground('#ffe0b2'); // Orange
    }
    
  } catch (error) {
    Logger.log(`Error updating initialization status: ${error}`);
  }
}

/**
 * Gets the session name from the YSLv6Hub sheet
 * @return {string} The session name or null if not set
 */
function getSessionName() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboardSheet = ss.getSheetByName('YSLv6Hub');
    
    if (!dashboardSheet) {
      return null;
    }
    
    // Get session name from row 4, column B
    const sessionName = dashboardSheet.getRange(4, 2).getValue();
    return sessionName || null;
    
  } catch (error) {
    Logger.log('Error getting session name: ' + error);
    return null;
  }
}

// Make functions available to other modules
const YSLv6Hub = {
  setupYSLv6Hub: setupYSLv6Hub,
  updateInitializationStatus: updateInitializationStatus,
  getSessionName: getSessionName
};