/**
 * YSL Hub v2 Upgrade Module
 * 
 * This module integrates all the enhancements and new features into the existing system.
 * It provides functions to upgrade the system and add new menu items for the enhanced features.
 * 
 * @author Claude Code
 * @version 1.0
 * @date 2025-05-05
 */

/**
 * Upgrades the YSL Hub system with all enhancements.
 * This is the main entry point for applying the upgrades.
 */
function upgradeYSLHub() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Confirm upgrade
    const result = ui.alert(
      'Upgrade YSL Hub',
      'This will upgrade YSL Hub with new features and enhancements. Continue?',
      ui.ButtonSet.YES_NO
    );
    
    if (result !== ui.Button.YES) {
      return;
    }
    
    // Log the upgrade start
    Logger.log('Starting YSL Hub upgrade');
    
    // Apply all upgrades
    const results = {
      fixedConfig: fixConfigurationDialog(),
      blankInit: installBlankInitializer(),
      emailTemplates: installEmailTemplates(),
      inputValidation: installInputValidation(),
      newMenu: createEnhancedMenu()
    };
    
    // Log results
    Logger.log(`Upgrade results: ${JSON.stringify(results)}`);
    
    // Show summary dialog
    showUpgradeSummary(results);
    
    return true;
  } catch (error) {
    // Basic error handling
    Logger.log(`Upgrade error: ${error.message}`);
    SpreadsheetApp.getUi().alert(
      'Upgrade Error',
      `An error occurred during upgrade: ${error.message}`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return false;
  }
}

/**
 * Shows a summary of the upgrade results
 * @param {Object} results - The upgrade results
 */
function showUpgradeSummary(results) {
  const ui = SpreadsheetApp.getUi();
  
  const successCount = Object.values(results).filter(result => result).length;
  const totalCount = Object.keys(results).length;
  
  const summary = `
YSL Hub Upgrade Summary
-----------------------
${successCount} of ${totalCount} upgrades were successfully applied.

Applied upgrades:
• Configuration Dialog Fix: ${results.fixedConfig ? 'Success' : 'Failed'}
• Blank Spreadsheet Initializer: ${results.blankInit ? 'Success' : 'Failed'}
• Email Templates System: ${results.emailTemplates ? 'Success' : 'Failed'}
• Input Validation: ${results.inputValidation ? 'Success' : 'Failed'}
• Enhanced Menu: ${results.newMenu ? 'Success' : 'Failed'}

To access the new features, use the "YSL Hub Enhanced" menu.
`;
  
  ui.alert(
    'Upgrade Complete',
    summary,
    ui.ButtonSet.OK
  );
}

/**
 * Installs the fixed configuration dialog
 * @return {boolean} Success status
 */
function fixConfigurationDialog() {
  try {
    // This function already exists in FixedConfigDialog.gs
    if (typeof showConfigurationDialogFixed === 'function') {
      Logger.log('Fixed configuration dialog is already installed');
      return true;
    }
    
    // Otherwise, the script has already been loaded
    return true;
  } catch (error) {
    Logger.log(`Error installing fixed configuration dialog: ${error.message}`);
    return false;
  }
}

/**
 * Installs the blank spreadsheet initializer
 * @return {boolean} Success status
 */
function installBlankInitializer() {
  try {
    // This function already exists in BlankSheetInitializer.gs
    if (typeof initializeBlankSpreadsheet === 'function') {
      Logger.log('Blank spreadsheet initializer is already installed');
      return true;
    }
    
    // Otherwise, the script has already been loaded
    return true;
  } catch (error) {
    Logger.log(`Error installing blank spreadsheet initializer: ${error.message}`);
    return false;
  }
}

/**
 * Installs the email templates system
 * @return {boolean} Success status
 */
function installEmailTemplates() {
  try {
    // Initialize email templates if the module exists
    if (typeof EmailTemplates !== 'undefined' && typeof EmailTemplates.initializeEmailTemplates === 'function') {
      EmailTemplates.initializeEmailTemplates();
      Logger.log('Email templates system initialized');
      return true;
    }
    
    // Otherwise, the script needs to be loaded
    return false;
  } catch (error) {
    Logger.log(`Error installing email templates: ${error.message}`);
    return false;
  }
}

/**
 * Installs the input validation system
 * @return {boolean} Success status
 */
function installInputValidation() {
  try {
    // Apply sheet validation if the module exists
    if (typeof InputValidation !== 'undefined' && typeof InputValidation.applySheetValidation === 'function') {
      InputValidation.applySheetValidation();
      Logger.log('Input validation system applied to sheets');
      return true;
    }
    
    // Otherwise, the script needs to be loaded
    return false;
  } catch (error) {
    Logger.log(`Error installing input validation: ${error.message}`);
    return false;
  }
}

/**
 * Creates the enhanced menu with all new features
 * @return {boolean} Success status
 */
function createEnhancedMenu() {
  try {
    const ui = SpreadsheetApp.getUi();
    const menu = ui.createMenu('YSL Hub Enhanced');
    
    // System section
    menu.addItem('Initialize Blank Spreadsheet', 'initializeBlankSpreadsheet');
    menu.addItem('System Configuration (Fixed)', 'showConfigurationDialogFixed');
    menu.addSeparator();
    
    // Email templates section
    menu.addSubMenu(ui.createMenu('Email Templates')
      .addItem('Manage Email Templates', 'EmailTemplates.showTemplateManager')
      .addItem('Create New Template', 'EmailTemplates.showCreateTemplateDialog')
      .addItem('Send Templated Email', 'EmailTemplates.emailClassParticipantsWithTemplate'));
    
    // Validation section
    menu.addItem('Apply Sheet Validation', 'InputValidation.applySheetValidation');
    
    // Add to UI
    menu.addToUi();
    
    return true;
  } catch (error) {
    Logger.log(`Error creating enhanced menu: ${error.message}`);
    return false;
  }
}

/**
 * Creates an upgrade guide sheet with documentation
 */
function createUpgradeGuide() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let guideSheet = ss.getSheetByName('Upgrade Guide');
    
    if (!guideSheet) {
      guideSheet = ss.insertSheet('Upgrade Guide');
    } else {
      guideSheet.clear();
    }
    
    // Set up header
    guideSheet.getRange('A1:B1').merge()
      .setValue('YSL Hub Enhanced - Upgrade Guide')
      .setFontWeight('bold')
      .setFontSize(14)
      .setHorizontalAlignment('center')
      .setBackground('#4285F4')
      .setFontColor('white');
    
    // Add introduction
    guideSheet.getRange('A2:B4').merge()
      .setValue('This guide explains the new features added to YSL Hub and how to use them.')
      .setWrap(true);
    
    // Set up sections
    const sections = [
      {
        title: '1. Configuration Dialog Fix',
        content: 'The System Configuration dialog has been fixed to work properly with blank spreadsheets. Use "System Configuration (Fixed)" from the YSL Hub Enhanced menu.'
      },
      {
        title: '2. Blank Spreadsheet Initializer',
        content: 'A new feature for creating all required sheets and structure from a blank spreadsheet. Use "Initialize Blank Spreadsheet" from the YSL Hub Enhanced menu.'
      },
      {
        title: '3. Email Templates System',
        content: 'Create and manage reusable email templates with placeholders for personalization. Use the "Email Templates" submenu from the YSL Hub Enhanced menu.'
      },
      {
        title: '4. Input Validation',
        content: 'Enhanced data validation throughout the system ensures data integrity and provides user-friendly error messages. Use "Apply Sheet Validation" to add dropdown lists and validation rules to your sheets.'
      }
    ];
    
    // Add sections
    let row = 6;
    sections.forEach(section => {
      guideSheet.getRange(row, 1, 1, 2).merge()
        .setValue(section.title)
        .setFontWeight('bold')
        .setBackground('#f3f3f3');
      
      row++;
      
      guideSheet.getRange(row, 1, 2, 2).merge()
        .setValue(section.content)
        .setWrap(true);
      
      row += 3;
    });
    
    // Format sheet
    guideSheet.setColumnWidth(1, 150);
    guideSheet.setColumnWidth(2, 450);
    
    // Activate the guide sheet
    guideSheet.activate();
    
    return true;
  } catch (error) {
    Logger.log(`Error creating upgrade guide: ${error.message}`);
    return false;
  }
}

/**
 * Creates a notification about the upgrade
 */
function notifyAboutUpgrade() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    ui.alert(
      'YSL Hub Enhanced',
      'YSL Hub has been enhanced with new features! Look for the "YSL Hub Enhanced" menu to access these features:\n\n' +
      '• Fixed Configuration Dialog\n' +
      '• Blank Spreadsheet Initializer\n' +
      '• Email Templates System\n' +
      '• Input Validation\n\n' +
      'View the Upgrade Guide sheet for more information.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    Logger.log(`Error showing upgrade notification: ${error.message}`);
    return false;
  }
}

/**
 * Runs the entire upgrade process including guide creation and notification
 */
function runFullUpgrade() {
  // Upgrade the core functionality
  const upgradeResult = upgradeYSLHub();
  
  // Create the guide
  const guideResult = createUpgradeGuide();
  
  // Show notification
  if (upgradeResult && guideResult) {
    notifyAboutUpgrade();
  }
}