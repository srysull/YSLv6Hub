/**
 * YSL Hub v2 Shared Placeholders Documentation
 * 
 * This file documents the placeholder constants used across modules.
 * It serves as documentation only - each module now has its own copy of 
 * the constants to prevent runtime dependency issues.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-18
 */

/**
 * IMPORTANT: This is a documentation file only!
 * 
 * The actual placeholder constants are defined locally in each module:
 * - 07_CommunicationModule.ts: COMMUNICATION_PLACEHOLDERS
 * - 08_ReportingModule.ts: REPORT_PLACEHOLDERS
 * 
 * The original approach of sharing a single SHARED_PLACEHOLDERS constant
 * caused runtime errors because Google Apps Script does not guarantee the
 * order of file execution, so the constant might not be available when needed.
 * 
 * When adding new placeholders, update them in both modules using this file
 * as a guide for consistency.
 */

// Documentation of all available placeholders for reference
const PLACEHOLDER_DOCUMENTATION = {
  // Common student/class placeholders
  STUDENT_NAME: '{{student_name}}',
  STUDENT_ID: '{{student_id}}',
  PARENT_NAME: '{{parent_name}}',
  CLASS_NAME: '{{class_name}}',
  INSTRUCTOR_NAME: '{{instructor_name}}',
  DAY: '{{day}}',
  TIME: '{{time}}',
  LOCATION: '{{location}}',
  START_DATE: '{{start_date}}',
  END_DATE: '{{end_date}}',
  SESSION_NAME: '{{session_name}}',
  LEVEL: '{{level}}',
  NEXT_LEVEL: '{{next_level}}',

  // Report-specific placeholders
  REPORT_DATE: '{{report_date}}',
  SKILLS_TABLE: '{{skills_table}}',
  COMMENTS: '{{comments}}',
  NEXT_STEPS: '{{next_steps}}',
  ASSESSMENT_SUMMARY: '{{assessment_summary}}',
  RECOMMENDED_CLASSES: '{{recommended_classes}}'
};