/**
 * YSLv6Hub Interfaces
 * 
 * This file exports all the interfaces used throughout the application.
 */

/**
 * Student interface representing a swim lesson student
 */
export interface Student {
  id: string;
  firstName: string;
  lastName: string;
  stage: string;
  age: number;
  guardian: string;
  email: string;
  phone: string;
  notes?: string;
}

/**
 * Class interface representing a swim lesson class
 */
export interface Class {
  id: string;
  name: string;
  stage: string;
  instructorName: string;
  schedule: string;
  location: string;
  startDate: Date;
  endDate: Date;
  studentIds: string[];
}

/**
 * Skill interface representing a swim skill
 */
export interface Skill {
  id: string;
  name: string;
  description: string;
  stage: string;
  category: string;
  order: number;
}

/**
 * SkillAssessment interface for tracking student skill progress
 */
export interface SkillAssessment {
  studentId: string;
  skillId: string;
  status: 'not-introduced' | 'introduced' | 'developing' | 'completed';
  assessedOn?: Date;
  notes?: string;
}

/**
 * Error context interface for structured error handling
 */
export interface ErrorContext {
  module: string;
  function: string;
  userMessage?: string;
  technicalDetails?: Record<string, any>;
}

/**
 * Cache options interface
 */
export interface CacheOptions {
  expirySeconds?: number;
}

/**
 * Log entry interface
 */
export interface LogEntry {
  timestamp: string;
  severity: string;
  module: string;
  function: string;
  message: string;
  details?: any;
  user?: string;
}

/**
 * SheetData interface for structured sheet data access
 */
export interface SheetData {
  values: any[][];
  headers: string[];
  namedRanges: Record<string, GoogleAppsScript.Spreadsheet.Range>;
}

/**
 * Template options interface
 */
export interface TemplateOptions {
  title?: string;
  width?: number;
  height?: number;
  modalMode?: boolean;
}

/**
 * Email template options
 */
export interface EmailOptions {
  to: string | string[];
  subject: string;
  cc?: string | string[];
  bcc?: string | string[];
  attachments?: GoogleAppsScript.Base.BlobSource[];
  name?: string;
  replyTo?: string;
  noReply?: boolean;
}

/**
 * Notification options
 */
export interface NotificationOptions {
  type: 'info' | 'warning' | 'error' | 'success';
  title?: string;
  timeout?: number;
}

/**
 * Sheet structure definition
 */
export interface SheetStructure {
  name: string;
  headers: string[];
  hiddenColumns?: number[];
  frozenRows?: number;
  frozenColumns?: number;
  columnWidths?: Record<number, number>;
  protectedRanges?: {
    range: string;
    description: string;
    editorsMode?: 'ONLY_USERS_WITH_PERMISSIONS' | 'ALL';
  }[];
}

/**
 * Menu item definition
 */
export interface MenuItem {
  label: string;
  functionName?: string;
  subMenu?: MenuItem[];
  separator?: boolean;
}