/**
 * YSL Hub v2 Shared Placeholders
 * 
 * This module provides shared placeholder constants for use across modules.
 * It prevents duplication of constants between modules and potential conflicts.
 * 
 * @author Sean R. Sullivan
 * @version 2.0
 * @date 2025-05-18
 */

// Shared placeholder constants used across modules
const SHARED_PLACEHOLDERS = {
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

// Global variable export
const PlaceholdersConstants = {
  SHARED_PLACEHOLDERS
};