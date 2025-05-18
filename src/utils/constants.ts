/**
 * YSLv6Hub Constants
 * 
 * This file contains global constants used throughout the application.
 */

/**
 * Sheet names
 */
export const SHEET_NAMES = {
  DASHBOARD: 'YSLv6Hub',
  REGISTRATION_INFO: 'RegistrationInfo',
  GROUPS_TRACKER: 'GroupsTracker',
  SWIMMER_SKILLS: 'SwimmerSkills',
  SWIMMER_LOG: 'SwimmerLog',
  COMMS_HUB: 'CommsHub',
  SYSTEM_LOG: 'SystemLog'
};

/**
 * Error severity levels
 */
export enum ErrorSeverity {
  INFO = 'info',
  WARNING = 'warning',
  ERROR = 'error',
  CRITICAL = 'critical'
}

/**
 * Feature flags for enabling/disabling features
 */
export enum FeatureFlag {
  SMART_IMPORT = 'SMART_IMPORT',
  FORM_PROCESSOR = 'FORM_PROCESSOR',
  ENHANCED_COMMUNICATIONS = 'ENHANCED_COMMUNICATIONS',
  ADVANCED_REPORTING = 'ADVANCED_REPORTING',
  SESSION_ANALYTICS = 'SESSION_ANALYTICS'
}

/**
 * User roles for access control
 */
export enum UserRole {
  ADMIN = 'ADMIN',
  COORDINATOR = 'COORDINATOR',
  INSTRUCTOR = 'INSTRUCTOR',
  VIEWER = 'VIEWER'
}

/**
 * User experience levels for progressive disclosure
 */
export enum UserExperienceLevel {
  BEGINNER = 'BEGINNER',
  INTERMEDIATE = 'INTERMEDIATE',
  ADVANCED = 'ADVANCED'
}

/**
 * Swim stage levels
 */
export enum SwimStage {
  STAGE_A = 'A',
  STAGE_B = 'B',
  STAGE_1 = '1',
  STAGE_2 = '2',
  STAGE_3 = '3',
  STAGE_4 = '4',
  STAGE_5 = '5',
  STAGE_6 = '6'
}

/**
 * Skill assessment status
 */
export enum SkillStatus {
  NOT_INTRODUCED = 'not-introduced',
  INTRODUCED = 'introduced',
  DEVELOPING = 'developing',
  COMPLETED = 'completed'
}

/**
 * Cell colors for skill assessments
 */
export const ASSESSMENT_COLORS = {
  NOT_INTRODUCED: '#ffffff', // White
  INTRODUCED: '#fff2cc',    // Light yellow
  DEVELOPING: '#ffe599',    // Darker yellow
  COMPLETED: '#b6d7a8'      // Green
};

/**
 * Column indices for GroupsTracker sheets
 */
export const GROUPS_TRACKER_COLUMNS = {
  STUDENT_NAME: 0,
  STUDENT_AGE: 1,
  // Skills start at column 2
};

/**
 * Column indices for SwimmerSkills sheets
 */
export const SWIMMER_SKILLS_COLUMNS = {
  STUDENT_NAME: 0,
  STUDENT_AGE: 1,
  CLASS_NAME: 2,
  // Skills start at column 3
};

/**
 * Script properties keys
 */
export const SCRIPT_PROPERTIES = {
  FEATURE_FLAGS: 'FEATURE_FLAGS',
  USER_ROLES: 'USER_ROLES',
  USER_EXPERIENCE_LEVELS: 'USER_EXPERIENCE_LEVELS',
  LAST_IMPORT_DATE: 'LAST_IMPORT_DATE',
  SYSTEM_VERSION: 'SYSTEM_VERSION'
};

/**
 * Cache keys
 */
export const CACHE_KEYS = {
  STUDENTS_PREFIX: 'students_',
  CLASSES_PREFIX: 'classes_',
  SKILLS_PREFIX: 'skills_',
  USER_PREFIX: 'user_'
};

/**
 * System version
 */
export const SYSTEM_VERSION = '6.0.0';

/**
 * Maximum log rows
 */
export const MAX_LOG_ROWS = 10000;

/**
 * Custom menu id
 */
export const MENU_ID = 'YSLv6Hub';

/**
 * Emergency menu id
 */
export const EMERGENCY_MENU_ID = 'YSLv6Hub Emergency';