/**
 * YSL Hub v2 Email Templates Module
 * 
 * This module provides a template system for emails sent through the YSL Hub.
 * It includes predefined templates, custom template storage, and template management.
 * 
 * @author Claude Code
 * @version 1.0
 * @date 2025-05-05
 */

// Template property key prefix
const TEMPLATE_PROPERTY_PREFIX = 'emailTemplate_';

// Default templates
const DEFAULT_TEMPLATES = {
  welcome: {
    name: 'Welcome Email',
    subject: 'Welcome to YSL Swim Lessons at PenBay YMCA',
    body: `<p>Dear {parent_name},</p>
<p>Welcome to the {session_name} session of Youth Swim Lessons at PenBay YMCA! We're excited to have {student_name} joining us for {program_name} on {class_day} at {class_time}.</p>

<p><strong>Important Information:</strong></p>
<ul>
  <li>Please arrive 10 minutes before your scheduled time</li>
  <li>Bring a swimsuit, towel, and goggles</li>
  <li>Meet your instructor, {instructor_name}, at the pool deck</li>
  <li>Parents are welcome to observe from the viewing area</li>
</ul>

<p>A copy of our Parent Handbook is attached to this email. It contains important information about our swim lesson policies.</p>

<p>If you have any questions, please don't hesitate to reach out.</p>

<p>Best regards,<br>
PenBay YMCA Aquatics Department<br>
ssullivan@penbayymca.org</p>`
  },
  
  lessonReminder: {
    name: 'Lesson Reminder',
    subject: 'Reminder: Your YSL Swim Lesson Tomorrow',
    body: `<p>Dear {parent_name},</p>
<p>This is a friendly reminder that {student_name} has swim lessons tomorrow, {class_day}, at {class_time}.</p>

<p><strong>Quick Reminders:</strong></p>
<ul>
  <li>Please arrive 10 minutes early</li>
  <li>Bring swimsuit, towel, and goggles</li>
  <li>Meet at the pool deck</li>
</ul>

<p>Looking forward to seeing you tomorrow!</p>

<p>Best regards,<br>
{instructor_name}<br>
PenBay YMCA Aquatics Department</p>`
  },
  
  assessmentComplete: {
    name: 'Assessment Completed',
    subject: 'YSL Swim Lesson Progress Report',
    body: `<p>Dear {parent_name},</p>
<p>We're pleased to share that {student_name} has completed their mid-session assessment for the {session_name} swim lessons.</p>

<p>Your child has been working on the following skills:</p>
<ul>
  <li>{skill_1}</li>
  <li>{skill_2}</li>
  <li>{skill_3}</li>
</ul>

<p>{instructor_notes}</p>

<p>A detailed progress report is attached to this email. Please review it and let us know if you have any questions.</p>

<p>Best regards,<br>
{instructor_name}<br>
PenBay YMCA Aquatics Department</p>`
  },
  
  classCancellation: {
    name: 'Class Cancellation',
    subject: 'IMPORTANT: Swim Lesson Cancelled',
    body: `<p>Dear {parent_name},</p>
<p><strong>IMPORTANT NOTICE:</strong> The swim lesson scheduled for {class_day}, {class_date} at {class_time} has been cancelled due to {cancellation_reason}.</p>

<p>We apologize for any inconvenience this may cause. The class will be rescheduled for {makeup_date} at {makeup_time}.</p>

<p>If you have any questions or concerns, please contact us immediately.</p>

<p>Best regards,<br>
PenBay YMCA Aquatics Department<br>
ssullivan@penbayymca.org</p>`
  },
  
  sessionEndingSummary: {
    name: 'Session Ending Summary',
    subject: 'YSL Session Wrap-Up and Next Steps',
    body: `<p>Dear {parent_name},</p>
<p>As we approach the end of our {session_name} swim lessons, we wanted to take a moment to thank you for your participation. {student_name} has shown great progress in their swimming skills!</p>

<p><strong>Session Achievements:</strong></p>
<ul>
  <li>{achievement_1}</li>
  <li>{achievement_2}</li>
  <li>{achievement_3}</li>
</ul>

<p>A final assessment report will be provided during the last class. Your instructor recommends {next_level_recommendation} for the next session.</p>

<p>Registration for the next session is now open. You can register at the front desk or online at www.penbayymca.org.</p>

<p>Thank you for choosing PenBay YMCA for your child's swim lessons!</p>

<p>Best regards,<br>
PenBay YMCA Aquatics Department<br>
ssullivan@penbayymca.org</p>`
  }
};

/**
 * Initializes the email template system
 * Ensures that default templates are available
 */
function initializeEmailTemplates() {
  try {
    // Log initialization
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Initializing email templates', 'INFO', 'initializeEmailTemplates');
    }
    
    // Store default templates if they don't exist
    for (const [id, template] of Object.entries(DEFAULT_TEMPLATES)) {
      const propertyKey = `${TEMPLATE_PROPERTY_PREFIX}${id}`;
      const existingTemplate = PropertiesService.getScriptProperties().getProperty(propertyKey);
      
      if (!existingTemplate) {
        const templateJson = JSON.stringify(template);
        PropertiesService.getScriptProperties().setProperty(propertyKey, templateJson);
        
        // Log creation of default template
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`Created default template: ${template.name}`, 'INFO', 'initializeEmailTemplates');
        }
      }
    }
    
    return true;
  } catch (error) {
    // Handle initialization errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'initializeEmailTemplates', 
        'Error initializing email templates. The system will use built-in templates as fallback.');
    } else {
      Logger.log(`Email template initialization error: ${error.message}`);
    }
    return false;
  }
}

/**
 * Gets all available email templates
 * @return {Object[]} Array of template objects
 */
function getAllTemplates() {
  try {
    const templates = [];
    const scriptProperties = PropertiesService.getScriptProperties().getProperties();
    
    // Get all template properties
    for (const key in scriptProperties) {
      if (key.startsWith(TEMPLATE_PROPERTY_PREFIX)) {
        try {
          const template = JSON.parse(scriptProperties[key]);
          const templateId = key.substring(TEMPLATE_PROPERTY_PREFIX.length);
          templates.push({
            id: templateId,
            name: template.name,
            subject: template.subject,
            body: template.body
          });
        } catch (parseError) {
          // Log parse error but continue processing other templates
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Error parsing template ${key}: ${parseError.message}`, 'ERROR', 'getAllTemplates');
          } else {
            Logger.log(`Error parsing template ${key}: ${parseError.message}`);
          }
        }
      }
    }
    
    // Sort templates by name
    templates.sort((a, b) => a.name.localeCompare(b.name));
    
    return templates;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting templates: ${error.message}`, 'ERROR', 'getAllTemplates');
    } else {
      Logger.log(`Error getting templates: ${error.message}`);
    }
    
    // Return default templates as fallback
    return Object.entries(DEFAULT_TEMPLATES).map(([id, template]) => ({
      id,
      name: template.name,
      subject: template.subject,
      body: template.body
    }));
  }
}

/**
 * Gets a specific template by ID
 * @param {string} templateId - The template ID
 * @return {Object|null} The template object or null if not found
 */
function getTemplate(templateId) {
  try {
    const propertyKey = `${TEMPLATE_PROPERTY_PREFIX}${templateId}`;
    const templateJson = PropertiesService.getScriptProperties().getProperty(propertyKey);
    
    if (templateJson) {
      const template = JSON.parse(templateJson);
      return {
        id: templateId,
        name: template.name,
        subject: template.subject,
        body: template.body
      };
    }
    
    // Check if it's a default template that hasn't been saved yet
    if (DEFAULT_TEMPLATES[templateId]) {
      return {
        id: templateId,
        name: DEFAULT_TEMPLATES[templateId].name,
        subject: DEFAULT_TEMPLATES[templateId].subject,
        body: DEFAULT_TEMPLATES[templateId].body
      };
    }
    
    return null;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error getting template ${templateId}: ${error.message}`, 'ERROR', 'getTemplate');
    } else {
      Logger.log(`Error getting template ${templateId}: ${error.message}`);
    }
    return null;
  }
}

/**
 * Saves a template
 * @param {string} templateId - The template ID
 * @param {string} name - The template name
 * @param {string} subject - The email subject
 * @param {string} body - The email body
 * @return {boolean} Success status
 */
function saveTemplate(templateId, name, subject, body) {
  try {
    // Validate inputs
    if (!templateId || !name || !subject || !body) {
      throw new Error('All template fields are required');
    }
    
    // Create template object
    const template = {
      name: name,
      subject: subject,
      body: body
    };
    
    // Save to script properties
    const propertyKey = `${TEMPLATE_PROPERTY_PREFIX}${templateId}`;
    const templateJson = JSON.stringify(template);
    PropertiesService.getScriptProperties().setProperty(propertyKey, templateJson);
    
    // Log template save
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Saved template: ${name} (${templateId})`, 'INFO', 'saveTemplate');
    }
    
    return true;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'saveTemplate', 
        'Error saving email template. Please try again or contact support.');
    } else {
      Logger.log(`Error saving template: ${error.message}`);
    }
    return false;
  }
}

/**
 * Deletes a template
 * @param {string} templateId - The template ID
 * @return {boolean} Success status
 */
function deleteTemplate(templateId) {
  try {
    // Don't allow deletion of default templates
    if (DEFAULT_TEMPLATES[templateId]) {
      throw new Error('Cannot delete default templates');
    }
    
    // Delete from script properties
    const propertyKey = `${TEMPLATE_PROPERTY_PREFIX}${templateId}`;
    PropertiesService.getScriptProperties().deleteProperty(propertyKey);
    
    // Log template deletion
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Deleted template: ${templateId}`, 'INFO', 'deleteTemplate');
    }
    
    return true;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'deleteTemplate', 
        'Error deleting email template. Please try again or contact support.');
    } else {
      Logger.log(`Error deleting template: ${error.message}`);
    }
    return false;
  }
}

/**
 * Fills a template with data
 * @param {string} templateId - The template ID
 * @param {Object} data - Key-value pairs for template variables
 * @return {Object} The filled template with subject and body
 */
function fillTemplate(templateId, data) {
  try {
    // Get the template
    const template = getTemplate(templateId);
    
    if (!template) {
      throw new Error(`Template not found: ${templateId}`);
    }
    
    // Fill in the template variables
    let subject = template.subject;
    let body = template.body;
    
    // Replace variables in the template
    for (const [key, value] of Object.entries(data)) {
      const placeholder = `{${key}}`;
      subject = subject.replace(new RegExp(placeholder, 'g'), value || '');
      body = body.replace(new RegExp(placeholder, 'g'), value || '');
    }
    
    // Remove any remaining placeholders
    subject = subject.replace(/{[a-z_]+}/g, '');
    body = body.replace(/{[a-z_]+}/g, '');
    
    return {
      subject: subject,
      body: body
    };
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage(`Error filling template ${templateId}: ${error.message}`, 'ERROR', 'fillTemplate');
    } else {
      Logger.log(`Error filling template: ${error.message}`);
    }
    
    // Return original template as fallback
    const template = getTemplate(templateId) || DEFAULT_TEMPLATES.welcome;
    return {
      subject: template.subject,
      body: template.body
    };
  }
}

/**
 * Shows the template manager UI
 */
function showTemplateManager() {
  try {
    const ui = SpreadsheetApp.getUi();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create or get the Templates sheet
    let templatesSheet = ss.getSheetByName('Email Templates');
    if (!templatesSheet) {
      templatesSheet = ss.insertSheet('Email Templates');
      
      // Set up headers
      const headers = ['Template ID', 'Template Name', 'Subject', 'Body', 'Actions'];
      templatesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      
      // Format header row
      templatesSheet.getRange(1, 1, 1, headers.length)
        .setFontWeight('bold')
        .setBackground('#f3f3f3')
        .setHorizontalAlignment('center');
      
      // Set column widths
      templatesSheet.setColumnWidth(1, 120);
      templatesSheet.setColumnWidth(2, 150);
      templatesSheet.setColumnWidth(3, 250);
      templatesSheet.setColumnWidth(4, 400);
      templatesSheet.setColumnWidth(5, 100);
      
      // Freeze header row
      templatesSheet.setFrozenRows(1);
    } else {
      // Clear existing template data but keep header
      const lastRow = templatesSheet.getLastRow();
      if (lastRow > 1) {
        templatesSheet.getRange(2, 1, lastRow - 1, 5).clear();
      }
    }
    
    // Get all templates
    const templates = getAllTemplates();
    
    // Add templates to sheet
    if (templates.length > 0) {
      const templateData = templates.map((template, index) => [
        template.id,
        template.name,
        template.subject,
        template.body,
        'Edit / Use'
      ]);
      
      templatesSheet.getRange(2, 1, templateData.length, 5).setValues(templateData);
      
      // Format body column for readability
      templatesSheet.getRange(2, 4, templateData.length, 1).setWrap(true);
    }
    
    // Add instructions
    templatesSheet.getRange(templates.length + 2, 1, 1, 5).merge()
      .setValue('Instructions: To edit a template, click "Edit" in the Actions column. To create a new template, click the "Create New Template" button below.')
      .setFontStyle('italic')
      .setHorizontalAlignment('center');
    
    // Add button cell for creating new template
    const newButtonCell = templatesSheet.getRange(templates.length + 4, 3);
    newButtonCell.setValue('Create New Template')
      .setBackground('#4285F4')
      .setFontColor('#FFFFFF')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    templatesSheet.setRowHeight(templates.length + 4, 30);
    
    // Activate the sheet
    templatesSheet.activate();
    
    // Show instructions
    ui.alert(
      'Email Template Manager',
      'You can view and manage email templates here. To use a template:\n\n' +
      '1. Find the template you want to use\n' +
      '2. Click "Use" in the Actions column\n' +
      '3. Fill in the required information\n\n' +
      'To create a new template, click the "Create New Template" button at the bottom of the sheet.',
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'showTemplateManager', 
        'Error showing template manager. Please try again or contact support.');
    } else {
      Logger.log(`Error showing template manager: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error showing template manager: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Shows the dialog to create a new template
 */
function showCreateTemplateDialog() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Get template ID
    const idResult = ui.prompt(
      'Create Email Template',
      'Enter a unique template ID (letters, numbers, and underscores only):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (idResult.getSelectedButton() !== ui.Button.OK) return false;
    const templateId = idResult.getResponseText().trim().replace(/[^a-z0-9_]/gi, '_');
    
    if (!templateId) {
      ui.alert('Error', 'Template ID is required.', ui.ButtonSet.OK);
      return false;
    }
    
    // Check if template already exists
    const existingTemplate = getTemplate(templateId);
    if (existingTemplate) {
      ui.alert('Error', `Template with ID "${templateId}" already exists.`, ui.ButtonSet.OK);
      return false;
    }
    
    // Get template name
    const nameResult = ui.prompt(
      'Create Email Template',
      'Enter a descriptive name for this template:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (nameResult.getSelectedButton() !== ui.Button.OK) return false;
    const name = nameResult.getResponseText().trim();
    
    if (!name) {
      ui.alert('Error', 'Template name is required.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get subject
    const subjectResult = ui.prompt(
      'Create Email Template',
      'Enter the email subject (you can use {placeholders}):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (subjectResult.getSelectedButton() !== ui.Button.OK) return false;
    const subject = subjectResult.getResponseText().trim();
    
    if (!subject) {
      ui.alert('Error', 'Email subject is required.', ui.ButtonSet.OK);
      return false;
    }
    
    // Get body (simplified version - ideally would use HTML dialog)
    const bodyResult = ui.prompt(
      'Create Email Template',
      'Enter the email body (you can use HTML and {placeholders}):',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (bodyResult.getSelectedButton() !== ui.Button.OK) return false;
    const body = bodyResult.getResponseText().trim();
    
    if (!body) {
      ui.alert('Error', 'Email body is required.', ui.ButtonSet.OK);
      return false;
    }
    
    // Save the template
    if (saveTemplate(templateId, name, subject, body)) {
      ui.alert(
        'Template Created',
        `The email template "${name}" has been created successfully.`,
        ui.ButtonSet.OK
      );
      
      // Refresh template manager
      showTemplateManager();
      return true;
    } else {
      ui.alert(
        'Error',
        'Failed to create template. Please try again.',
        ui.ButtonSet.OK
      );
      return false;
    }
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'showCreateTemplateDialog', 
        'Error creating email template. Please try again or contact support.');
    } else {
      Logger.log(`Error creating template: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error creating template: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Handles template edit
 * @param {string} templateId - The template ID to edit
 */
function editTemplate(templateId) {
  try {
    const ui = SpreadsheetApp.getUi();
    const template = getTemplate(templateId);
    
    if (!template) {
      ui.alert('Error', `Template not found: ${templateId}`, ui.ButtonSet.OK);
      return false;
    }
    
    // Get updated name
    const nameResult = ui.prompt(
      'Edit Email Template',
      'Edit template name:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (nameResult.getSelectedButton() !== ui.Button.OK) return false;
    const name = nameResult.getResponseText().trim() || template.name;
    
    // Get updated subject
    const subjectResult = ui.prompt(
      'Edit Email Template',
      'Edit email subject:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (subjectResult.getSelectedButton() !== ui.Button.OK) return false;
    const subject = subjectResult.getResponseText().trim() || template.subject;
    
    // Get updated body
    const bodyResult = ui.prompt(
      'Edit Email Template',
      'Edit email body:',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (bodyResult.getSelectedButton() !== ui.Button.OK) return false;
    const body = bodyResult.getResponseText().trim() || template.body;
    
    // Save the updated template
    if (saveTemplate(templateId, name, subject, body)) {
      ui.alert(
        'Template Updated',
        `The email template "${name}" has been updated successfully.`,
        ui.ButtonSet.OK
      );
      
      // Refresh template manager
      showTemplateManager();
      return true;
    } else {
      ui.alert(
        'Error',
        'Failed to update template. Please try again.',
        ui.ButtonSet.OK
      );
      return false;
    }
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'editTemplate', 
        'Error editing email template. Please try again or contact support.');
    } else {
      Logger.log(`Error editing template: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error editing template: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

/**
 * Enhanced version of emailClassParticipants that uses templates
 */
function emailClassParticipantsWithTemplate() {
  try {
    const ui = SpreadsheetApp.getUi();
    
    // Log operation start
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Starting email to class participants with template', 'INFO', 'emailClassParticipantsWithTemplate');
    }
    
    // Initialize templates if needed
    initializeEmailTemplates();
    
    // Get selected classes
    const selectedClasses = CommunicationModule.getSelectedClasses();
    if (selectedClasses.length === 0) {
      ui.alert(
        'No Classes Selected',
        'Please select at least one class in the Classes sheet by setting the "Select Class" column to "Select".',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Get templates
    const templates = getAllTemplates();
    if (templates.length === 0) {
      ui.alert(
        'No Templates Available',
        'No email templates are available. Please create templates first.',
        ui.ButtonSet.OK
      );
      return false;
    }
    
    // Create template selection menu
    const templateOptions = templates.map(t => t.name).join('|');
    const templateResult = ui.prompt(
      'Select Email Template',
      `Choose a template for this email (available templates: ${templateOptions}):`,
      ui.ButtonSet.OK_CANCEL
    );
    
    if (templateResult.getSelectedButton() !== ui.Button.OK) return false;
    const templateName = templateResult.getResponseText().trim();
    
    // Find the template by name
    const template = templates.find(t => t.name.toLowerCase() === templateName.toLowerCase());
    if (!template) {
      ui.alert('Error', `Template not found: ${templateName}`, ui.ButtonSet.OK);
      return false;
    }
    
    // Confirm sending
    const totalStudents = selectedClasses.reduce((sum, classInfo) => sum + classInfo.count, 0);
    
    let confirmed = false;
    if (ErrorHandling && typeof ErrorHandling.confirmAction === 'function') {
      confirmed = ErrorHandling.confirmAction(
        'Confirm Email',
        `This will send the "${template.name}" template to approximately ${totalStudents} participants across ${selectedClasses.length} classes. Continue?`
      );
    } else {
      // Fallback confirmation
      const result = ui.alert(
        'Confirm Email',
        `This will send the "${template.name}" template to approximately ${totalStudents} participants across ${selectedClasses.length} classes. Continue?`,
        ui.ButtonSet.YES_NO
      );
      
      confirmed = (result === ui.Button.YES);
    }
    
    if (!confirmed) return false;
    
    // Process each class
    let emailsSent = 0;
    let emailsFailed = 0;
    
    for (const classInfo of selectedClasses) {
      // Get class roster
      const roster = CommunicationModule.getClassRoster(classInfo.id);
      
      if (!roster || roster.length === 0) {
        if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
          ErrorHandling.logMessage(`No students found for class: ${classInfo.program} (${classInfo.day}, ${classInfo.time})`, 'WARNING', 'emailClassParticipantsWithTemplate');
        }
        continue;
      }
      
      // Process each student
      for (const student of roster) {
        try {
          // Skip if no email
          if (!student.email) {
            if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
              ErrorHandling.logMessage(`No email found for student: ${student.name}`, 'WARNING', 'emailClassParticipantsWithTemplate');
            }
            continue;
          }
          
          // Prepare template data
          const templateData = {
            student_name: student.name,
            parent_name: student.parent || student.name.split(' ')[0] + '\'s Parent/Guardian',
            program_name: classInfo.program,
            class_day: classInfo.day,
            class_time: classInfo.time,
            instructor_name: classInfo.instructor,
            session_name: GlobalFunctions.safeGetProperty(CONFIG.SESSION_NAME) || 'current session'
          };
          
          // Fill template
          const filledTemplate = fillTemplate(template.id, templateData);
          
          // Send email
          MailApp.sendEmail({
            to: student.email,
            subject: filledTemplate.subject,
            htmlBody: filledTemplate.body,
            name: DEFAULT_SENDER.NAME
          });
          
          emailsSent++;
          
          // Log success
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Email sent to ${student.email} for student ${student.name}`, 'INFO', 'emailClassParticipantsWithTemplate');
          }
        } catch (emailError) {
          emailsFailed++;
          
          // Log error
          if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
            ErrorHandling.logMessage(`Failed to send email to ${student.email}: ${emailError.message}`, 'ERROR', 'emailClassParticipantsWithTemplate');
          } else {
            Logger.log(`Failed to send email to ${student.email}: ${emailError.message}`);
          }
        }
      }
    }
    
    // Show results
    ui.alert(
      'Email Results',
      `Emails sent: ${emailsSent}\nEmails failed: ${emailsFailed}`,
      ui.ButtonSet.OK
    );
    
    return true;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'emailClassParticipantsWithTemplate', 
        'Error sending emails with template. Please try again or contact support.');
    } else {
      Logger.log(`Error sending emails with template: ${error.message}`);
      SpreadsheetApp.getUi().alert('Error', 
        `Error sending emails with template: ${error.message}`, 
        SpreadsheetApp.getUi().ButtonSet.OK);
    }
    return false;
  }
}

// Make functions available to other modules
const EmailTemplates = {
  initializeEmailTemplates: initializeEmailTemplates,
  getAllTemplates: getAllTemplates,
  getTemplate: getTemplate,
  saveTemplate: saveTemplate,
  deleteTemplate: deleteTemplate,
  fillTemplate: fillTemplate,
  showTemplateManager: showTemplateManager,
  showCreateTemplateDialog: showCreateTemplateDialog,
  editTemplate: editTemplate,
  emailClassParticipantsWithTemplate: emailClassParticipantsWithTemplate
};