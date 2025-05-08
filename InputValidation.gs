/**
 * YSL Hub v2 Input Validation Module
 * 
 * This module provides validation functions for user inputs across the system.
 * It helps ensure data integrity and provides user-friendly error messages.
 * 
 * @author Claude Code
 * @version 1.0
 * @date 2025-05-05
 */

// Validation error types
const VALIDATION_ERROR = {
  REQUIRED: 'required',
  FORMAT: 'format',
  RANGE: 'range',
  EXISTS: 'exists',
  UNIQUE: 'unique',
  DEPENDENCY: 'dependency'
};

// Common validation patterns
const VALIDATION_PATTERNS = {
  EMAIL: /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/,
  PHONE: /^(\+\d{1,2}\s?)?\(?\d{3}\)?[\s.-]?\d{3}[\s.-]?\d{4}$/,
  URL: /^(https?:\/\/)?([\da-z.-]+)\.([a-z.]{2,6})([\/\w .-]*)*\/?$/,
  DATE: /^\d{4}[-/](0?[1-9]|1[012])[-/](0?[1-9]|[12][0-9]|3[01])$/,
  TIME: /^(0?[1-9]|1[0-2]):[0-5][0-9]\s?([AP]M)?$/i,
  GOOGLE_DRIVE_URL: /^https:\/\/drive\.google\.com\/(file\/d\/|drive\/folders\/|spreadsheets\/d\/)([a-zA-Z0-9_-]+)/
};

/**
 * Validates a required field
 * @param {*} value - The value to validate
 * @param {string} fieldName - The name of the field for error messages
 * @return {Object} Validation result with success status and error message
 */
function validateRequired(value, fieldName) {
  const isValid = value !== undefined && value !== null && value !== '';
  
  return {
    valid: isValid,
    error: isValid ? '' : `${fieldName} is required`,
    errorType: isValid ? '' : VALIDATION_ERROR.REQUIRED
  };
}

/**
 * Validates a value against a regex pattern
 * @param {string} value - The value to validate
 * @param {RegExp} pattern - The regex pattern
 * @param {string} fieldName - The name of the field for error messages
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validatePattern(value, pattern, fieldName, required = true) {
  // If not required and empty, return valid
  if (!required && (value === undefined || value === null || value === '')) {
    return {
      valid: true,
      error: '',
      errorType: ''
    };
  }
  
  // Check required first
  const requiredCheck = validateRequired(value, fieldName);
  if (!requiredCheck.valid) {
    return requiredCheck;
  }
  
  // Check pattern
  const isValid = pattern.test(value);
  
  return {
    valid: isValid,
    error: isValid ? '' : `${fieldName} has an invalid format`,
    errorType: isValid ? '' : VALIDATION_ERROR.FORMAT
  };
}

/**
 * Validates an email address
 * @param {string} email - The email to validate
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateEmail(email, required = true) {
  return validatePattern(
    email, 
    VALIDATION_PATTERNS.EMAIL, 
    'Email address', 
    required
  );
}

/**
 * Validates a phone number
 * @param {string} phone - The phone number to validate
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validatePhone(phone, required = true) {
  return validatePattern(
    phone, 
    VALIDATION_PATTERNS.PHONE, 
    'Phone number', 
    required
  );
}

/**
 * Validates a URL
 * @param {string} url - The URL to validate
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateUrl(url, required = true) {
  return validatePattern(
    url, 
    VALIDATION_PATTERNS.URL, 
    'URL', 
    required
  );
}

/**
 * Validates a Google Drive URL
 * @param {string} url - The Google Drive URL to validate
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateGoogleDriveUrl(url, required = true) {
  return validatePattern(
    url, 
    VALIDATION_PATTERNS.GOOGLE_DRIVE_URL, 
    'Google Drive URL', 
    required
  );
}

/**
 * Validates a date string
 * @param {string} date - The date string to validate (YYYY-MM-DD or YYYY/MM/DD)
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateDate(date, required = true) {
  return validatePattern(
    date, 
    VALIDATION_PATTERNS.DATE, 
    'Date', 
    required
  );
}

/**
 * Validates a time string
 * @param {string} time - The time string to validate (HH:MM AM/PM)
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateTime(time, required = true) {
  return validatePattern(
    time, 
    VALIDATION_PATTERNS.TIME, 
    'Time', 
    required
  );
}

/**
 * Validates a numeric value within a range
 * @param {number} value - The value to validate
 * @param {number} min - The minimum allowed value
 * @param {number} max - The maximum allowed value
 * @param {string} fieldName - The name of the field for error messages
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateNumberRange(value, min, max, fieldName, required = true) {
  // If not required and empty, return valid
  if (!required && (value === undefined || value === null || value === '')) {
    return {
      valid: true,
      error: '',
      errorType: ''
    };
  }
  
  // Check required first
  const requiredCheck = validateRequired(value, fieldName);
  if (!requiredCheck.valid) {
    return requiredCheck;
  }
  
  // Convert to number if string
  const numValue = typeof value === 'string' ? parseFloat(value) : value;
  
  // Check if it's a valid number
  if (isNaN(numValue)) {
    return {
      valid: false,
      error: `${fieldName} must be a valid number`,
      errorType: VALIDATION_ERROR.FORMAT
    };
  }
  
  // Check range
  const isValid = numValue >= min && numValue <= max;
  
  return {
    valid: isValid,
    error: isValid ? '' : `${fieldName} must be between ${min} and ${max}`,
    errorType: isValid ? '' : VALIDATION_ERROR.RANGE
  };
}

/**
 * Validates a text length
 * @param {string} text - The text to validate
 * @param {number} minLength - The minimum allowed length
 * @param {number} maxLength - The maximum allowed length
 * @param {string} fieldName - The name of the field for error messages
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateTextLength(text, minLength, maxLength, fieldName, required = true) {
  // If not required and empty, return valid
  if (!required && (text === undefined || text === null || text === '')) {
    return {
      valid: true,
      error: '',
      errorType: ''
    };
  }
  
  // Check required first
  const requiredCheck = validateRequired(text, fieldName);
  if (!requiredCheck.valid) {
    return requiredCheck;
  }
  
  // Check text length
  const length = text.length;
  const isValid = length >= minLength && length <= maxLength;
  
  return {
    valid: isValid,
    error: isValid ? '' : `${fieldName} must be between ${minLength} and ${maxLength} characters`,
    errorType: isValid ? '' : VALIDATION_ERROR.RANGE
  };
}

/**
 * Validates that a value exists in a list of options
 * @param {*} value - The value to validate
 * @param {Array} options - The list of allowed options
 * @param {string} fieldName - The name of the field for error messages
 * @param {boolean} required - Whether the field is required
 * @return {Object} Validation result with success status and error message
 */
function validateOption(value, options, fieldName, required = true) {
  // If not required and empty, return valid
  if (!required && (value === undefined || value === null || value === '')) {
    return {
      valid: true,
      error: '',
      errorType: ''
    };
  }
  
  // Check required first
  const requiredCheck = validateRequired(value, fieldName);
  if (!requiredCheck.valid) {
    return requiredCheck;
  }
  
  // Check if value exists in options
  const isValid = options.includes(value);
  
  return {
    valid: isValid,
    error: isValid ? '' : `${fieldName} must be one of: ${options.join(', ')}`,
    errorType: isValid ? '' : VALIDATION_ERROR.EXISTS
  };
}

/**
 * Validates configuration input
 * @param {Object} config - The configuration object to validate
 * @return {Object} Validation result with success status, error messages, and valid config
 */
function validateConfiguration(config) {
  const errors = [];
  const validatedConfig = {};
  
  // Validate session name
  const sessionNameResult = validateRequired(config.sessionName, 'Session Name');
  if (!sessionNameResult.valid) {
    errors.push(sessionNameResult.error);
  } else {
    validatedConfig.sessionName = config.sessionName.trim();
  }
  
  // Validate roster folder URL
  const rosterFolderResult = validateGoogleDriveUrl(config.rosterFolderUrl, true);
  if (!rosterFolderResult.valid) {
    errors.push(rosterFolderResult.error);
  } else {
    validatedConfig.rosterFolderUrl = config.rosterFolderUrl.trim();
  }
  
  // Validate report template URL
  const reportTemplateResult = validateGoogleDriveUrl(config.reportTemplateUrl, true);
  if (!reportTemplateResult.valid) {
    errors.push(reportTemplateResult.error);
  } else {
    validatedConfig.reportTemplateUrl = config.reportTemplateUrl.trim();
  }
  
  // Validate swimmer records URL
  const swimmerRecordsResult = validateGoogleDriveUrl(config.swimmerRecordsUrl, true);
  if (!swimmerRecordsResult.valid) {
    errors.push(swimmerRecordsResult.error);
  } else {
    validatedConfig.swimmerRecordsUrl = config.swimmerRecordsUrl.trim();
  }
  
  // Validate session programs URL
  const sessionProgramsResult = validateGoogleDriveUrl(config.sessionProgramsUrl, true);
  if (!sessionProgramsResult.valid) {
    errors.push(sessionProgramsResult.error);
  } else {
    validatedConfig.sessionProgramsUrl = config.sessionProgramsUrl.trim();
  }
  
  // Validate parent handbook URL (optional)
  if (config.parentHandbookUrl) {
    const parentHandbookResult = validateGoogleDriveUrl(config.parentHandbookUrl, false);
    if (!parentHandbookResult.valid) {
      errors.push(parentHandbookResult.error);
    } else {
      validatedConfig.parentHandbookUrl = config.parentHandbookUrl.trim();
    }
  } else {
    validatedConfig.parentHandbookUrl = '';
  }
  
  return {
    valid: errors.length === 0,
    errors: errors,
    config: validatedConfig
  };
}

/**
 * Validates a class definition
 * @param {Object} classInfo - The class information to validate
 * @return {Object} Validation result with success status, error messages, and valid classInfo
 */
function validateClass(classInfo) {
  const errors = [];
  const validatedClass = {};
  
  // Validate program
  const programResult = validateRequired(classInfo.program, 'Program');
  if (!programResult.valid) {
    errors.push(programResult.error);
  } else {
    validatedClass.program = classInfo.program.trim();
  }
  
  // Validate day
  const dayOptions = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'];
  const dayResult = validateOption(classInfo.day, dayOptions, 'Day');
  if (!dayResult.valid) {
    errors.push(dayResult.error);
  } else {
    validatedClass.day = classInfo.day;
  }
  
  // Validate time
  const timeResult = validateTime(classInfo.time);
  if (!timeResult.valid) {
    errors.push(timeResult.error);
  } else {
    validatedClass.time = classInfo.time;
  }
  
  // Validate location
  const locationResult = validateRequired(classInfo.location, 'Location');
  if (!locationResult.valid) {
    errors.push(locationResult.error);
  } else {
    validatedClass.location = classInfo.location.trim();
  }
  
  // Validate count
  const countResult = validateNumberRange(classInfo.count, 0, 100, 'Student Count');
  if (!countResult.valid) {
    errors.push(countResult.error);
  } else {
    validatedClass.count = parseInt(classInfo.count);
  }
  
  // Validate instructor (optional)
  if (classInfo.instructor) {
    validatedClass.instructor = classInfo.instructor.trim();
  } else {
    validatedClass.instructor = '';
  }
  
  return {
    valid: errors.length === 0,
    errors: errors,
    class: validatedClass
  };
}

/**
 * Validates a student record
 * @param {Object} student - The student information to validate
 * @return {Object} Validation result with success status, error messages, and valid student
 */
function validateStudent(student) {
  const errors = [];
  const validatedStudent = {};
  
  // Validate name
  const nameResult = validateRequired(student.name, 'Student Name');
  if (!nameResult.valid) {
    errors.push(nameResult.error);
  } else {
    validatedStudent.name = student.name.trim();
  }
  
  // Validate age
  const ageResult = validateNumberRange(student.age, 1, 120, 'Age');
  if (!ageResult.valid) {
    errors.push(ageResult.error);
  } else {
    validatedStudent.age = parseInt(student.age);
  }
  
  // Validate parent/guardian
  const parentResult = validateRequired(student.parent, 'Parent/Guardian Name');
  if (!parentResult.valid) {
    errors.push(parentResult.error);
  } else {
    validatedStudent.parent = student.parent.trim();
  }
  
  // Validate email
  const emailResult = validateEmail(student.email);
  if (!emailResult.valid) {
    errors.push(emailResult.error);
  } else {
    validatedStudent.email = student.email.trim();
  }
  
  // Validate phone
  const phoneResult = validatePhone(student.phone);
  if (!phoneResult.valid) {
    errors.push(phoneResult.error);
  } else {
    validatedStudent.phone = student.phone;
  }
  
  // Special notes (optional)
  if (student.notes) {
    validatedStudent.notes = student.notes.trim();
  } else {
    validatedStudent.notes = '';
  }
  
  return {
    valid: errors.length === 0,
    errors: errors,
    student: validatedStudent
  };
}

/**
 * Validates an announcement
 * @param {Object} announcement - The announcement to validate
 * @return {Object} Validation result with success status, error messages, and valid announcement
 */
function validateAnnouncement(announcement) {
  const errors = [];
  const validatedAnnouncement = {};
  
  // Validate class ID
  const classIdResult = validateRequired(announcement.classId, 'Class ID');
  if (!classIdResult.valid) {
    errors.push(classIdResult.error);
  } else {
    validatedAnnouncement.classId = announcement.classId.trim();
  }
  
  // Validate subject
  const subjectResult = validateTextLength(announcement.subject, 3, 100, 'Subject');
  if (!subjectResult.valid) {
    errors.push(subjectResult.error);
  } else {
    validatedAnnouncement.subject = announcement.subject.trim();
  }
  
  // Validate message
  const messageResult = validateTextLength(announcement.message, 5, 10000, 'Message');
  if (!messageResult.valid) {
    errors.push(messageResult.error);
  } else {
    validatedAnnouncement.message = announcement.message.trim();
  }
  
  // Validate status
  const statusOptions = ['Draft', 'Ready', 'Sent', 'Failed'];
  const statusResult = validateOption(announcement.status, statusOptions, 'Status');
  if (!statusResult.valid) {
    errors.push(statusResult.error);
  } else {
    validatedAnnouncement.status = announcement.status;
  }
  
  return {
    valid: errors.length === 0,
    errors: errors,
    announcement: validatedAnnouncement
  };
}

/**
 * Applies data validation rules to sheets
 * This adds dropdown lists and other validation rules to the sheets
 */
function applySheetValidation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Classes sheet validation
    const classesSheet = ss.getSheetByName('Classes');
    if (classesSheet) {
      // Select Class column
      const selectRange = classesSheet.getRange("A2:A100");
      const selectRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Select', 'Exclude'], true)
        .build();
      selectRange.setDataValidation(selectRule);
      
      // Day column
      const dayRange = classesSheet.getRange("C2:C100");
      const dayRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'], true)
        .build();
      dayRange.setDataValidation(dayRule);
      
      // Count column - numeric only
      const countRange = classesSheet.getRange("F2:F100");
      const countRule = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(0, 100)
        .build();
      countRange.setDataValidation(countRule);
    }
    
    // Announcements sheet validation
    const announcementsSheet = ss.getSheetByName('Announcements');
    if (announcementsSheet) {
      // Status column
      const statusRange = announcementsSheet.getRange("I2:I100");
      const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Draft', 'Ready', 'Sent', 'Failed'], true)
        .build();
      statusRange.setDataValidation(statusRule);
    }
    
    // Assessments sheet validation
    const assessmentsSheet = ss.getSheetByName('Assessments');
    if (assessmentsSheet) {
      // Rating column
      const ratingRange = assessmentsSheet.getRange("E2:E100");
      const ratingRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['Not Assessed', 'Not Yet', 'In Progress', 'Completed'], true)
        .build();
      ratingRange.setDataValidation(ratingRule);
    }
    
    // Log successful application
    if (ErrorHandling && typeof ErrorHandling.logMessage === 'function') {
      ErrorHandling.logMessage('Sheet validation rules applied successfully', 'INFO', 'applySheetValidation');
    }
    
    return true;
  } catch (error) {
    // Handle errors
    if (ErrorHandling && typeof ErrorHandling.handleError === 'function') {
      ErrorHandling.handleError(error, 'applySheetValidation', 
        'Error applying sheet validation rules. Please try again or contact support.');
    } else {
      Logger.log(`Error applying sheet validation: ${error.message}`);
    }
    return false;
  }
}

/**
 * Shows a validation error dialog with a list of errors
 * @param {string[]} errors - Array of error messages
 * @param {string} title - Dialog title
 */
function showValidationErrorDialog(errors, title = 'Validation Error') {
  try {
    const ui = SpreadsheetApp.getUi();
    
    const errorList = errors.map(error => `• ${error}`).join('\n');
    const message = `Please fix the following errors:\n\n${errorList}`;
    
    ui.alert(
      title,
      message,
      ui.ButtonSet.OK
    );
  } catch (error) {
    // Fallback to simple alert if custom dialog fails
    const errorList = errors.join(', ');
    SpreadsheetApp.getUi().alert(`Validation Error: ${errorList}`);
  }
}

/**
 * Enhanced version of configuration validation that presents a user-friendly dialog
 * @param {Object} config - The configuration object to validate
 * @return {Object} The validated configuration object or null if invalid
 */
function validateConfigurationWithFeedback(config) {
  // Validate the configuration
  const validationResult = validateConfiguration(config);
  
  // Show errors if any
  if (!validationResult.valid) {
    showValidationErrorDialog(validationResult.errors, 'Configuration Error');
    return null;
  }
  
  return validationResult.config;
}

// Make functions available to other modules
const InputValidation = {
  validateRequired: validateRequired,
  validateEmail: validateEmail,
  validatePhone: validatePhone,
  validateUrl: validateUrl,
  validateGoogleDriveUrl: validateGoogleDriveUrl,
  validateDate: validateDate,
  validateTime: validateTime,
  validateNumberRange: validateNumberRange,
  validateTextLength: validateTextLength,
  validateOption: validateOption,
  validateConfiguration: validateConfiguration,
  validateClass: validateClass,
  validateStudent: validateStudent,
  validateAnnouncement: validateAnnouncement,
  applySheetValidation: applySheetValidation,
  showValidationErrorDialog: showValidationErrorDialog,
  validateConfigurationWithFeedback: validateConfigurationWithFeedback
};