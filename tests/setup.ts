/**
 * Jest setup file for Google Apps Script mocking
 * This file sets up the global mocks for all GAS APIs used in tests
 */

// Mock SpreadsheetApp
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(() => ({
    getSheetByName: jest.fn(),
    insertSheet: jest.fn(),
    getSheets: jest.fn(() => []),
    getName: jest.fn(() => 'Test Spreadsheet'),
    getId: jest.fn(() => 'test-spreadsheet-id'),
    getUrl: jest.fn(() => 'https://docs.google.com/spreadsheets/d/test-spreadsheet-id'),
  })),
  getUi: jest.fn(() => ({
    alert: jest.fn(),
    prompt: jest.fn(),
    showModalDialog: jest.fn(),
    showSidebar: jest.fn(),
    createMenu: jest.fn(() => ({
      addItem: jest.fn().mockReturnThis(),
      addSeparator: jest.fn().mockReturnThis(),
      addSubMenu: jest.fn().mockReturnThis(),
      addToUi: jest.fn(),
    })),
    ButtonSet: {
      OK: 'OK',
      OK_CANCEL: 'OK_CANCEL',
      YES_NO: 'YES_NO',
    },
    Button: {
      OK: 'OK',
      CANCEL: 'CANCEL',
      YES: 'YES',
      NO: 'NO',
    },
  })),
  create: jest.fn(() => ({
    getSheetByName: jest.fn(),
    insertSheet: jest.fn(),
  })),
} as any;

// Mock HtmlService
global.HtmlService = {
  createHtmlOutputFromFile: jest.fn(() => ({
    setWidth: jest.fn().mockReturnThis(),
    setHeight: jest.fn().mockReturnThis(),
    setSandboxMode: jest.fn().mockReturnThis(),
    append: jest.fn().mockReturnThis(),
    getContent: jest.fn(() => '<html></html>'),
  })),
  createHtmlOutput: jest.fn((html: string) => ({
    setWidth: jest.fn().mockReturnThis(),
    setHeight: jest.fn().mockReturnThis(),
    setSandboxMode: jest.fn().mockReturnThis(),
    append: jest.fn().mockReturnThis(),
    getContent: jest.fn(() => html),
  })),
  SandboxMode: {
    IFRAME: 'IFRAME',
  },
} as any;

// Mock PropertiesService
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
  })),
  getUserProperties: jest.fn(() => ({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn(() => ({})),
    setProperties: jest.fn(),
  })),
} as any;

// Mock DriveApp
global.DriveApp = {
  getFileById: jest.fn(() => ({
    getName: jest.fn(() => 'Test File'),
    getBlob: jest.fn(),
    makeCopy: jest.fn(),
  })),
  getFolderById: jest.fn(() => ({
    getName: jest.fn(() => 'Test Folder'),
    createFile: jest.fn(),
    getFiles: jest.fn(() => ({
      hasNext: jest.fn(() => false),
      next: jest.fn(),
    })),
  })),
} as any;

// Mock MailApp
global.MailApp = {
  sendEmail: jest.fn(),
  getRemainingDailyQuota: jest.fn(() => 100),
} as any;

// Mock Utilities
global.Utilities = {
  formatDate: jest.fn((date: Date, _timeZone: string, _format: string) => {
    return date.toISOString();
  }),
  formatString: jest.fn((template: string, ...args: any[]) => {
    let result = template;
    args.forEach((arg: any, index: number) => {
      result = result.replace(`%${index + 1}$s`, String(arg));
    });
    return result;
  }),
  sleep: jest.fn(),
} as any;

// Mock Logger
global.Logger = {
  log: jest.fn(),
} as any;

// Mock console for GAS environment
global.console = {
  log: jest.fn(),
  error: jest.fn(),
  warn: jest.fn(),
  info: jest.fn(),
} as any;

// Export for use in tests
export const mockSpreadsheet = global.SpreadsheetApp.getActiveSpreadsheet();
export const mockUi = global.SpreadsheetApp.getUi();
export const mockProperties = global.PropertiesService.getScriptProperties();