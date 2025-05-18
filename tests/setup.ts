/**
 * YSLv6Hub Test Setup
 * 
 * This file configures the test environment and provides mocks for Google Apps Script services
 */

// Mock global Google Apps Script objects
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(),
  getUi: jest.fn().mockReturnValue({
    alert: jest.fn(),
    prompt: jest.fn(),
    ButtonSet: {
      OK: 'OK',
      OK_CANCEL: 'OK_CANCEL',
      YES_NO: 'YES_NO',
      YES_NO_CANCEL: 'YES_NO_CANCEL'
    }
  })
};

global.Session = {
  getActiveUser: jest.fn().mockReturnValue({
    getEmail: jest.fn().mockReturnValue('test@example.com')
  }),
  getEffectiveUser: jest.fn().mockReturnValue({
    getEmail: jest.fn().mockReturnValue('test@example.com')
  })
};

global.PropertiesService = {
  getScriptProperties: jest.fn().mockReturnValue({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn().mockReturnValue({})
  }),
  getUserProperties: jest.fn().mockReturnValue({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn().mockReturnValue({})
  }),
  getDocumentProperties: jest.fn().mockReturnValue({
    getProperty: jest.fn(),
    setProperty: jest.fn(),
    deleteProperty: jest.fn(),
    getProperties: jest.fn().mockReturnValue({})
  })
};

global.CacheService = {
  getScriptCache: jest.fn().mockReturnValue({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn()
  }),
  getUserCache: jest.fn().mockReturnValue({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn()
  }),
  getDocumentCache: jest.fn().mockReturnValue({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn()
  })
};

global.HtmlService = {
  createHtmlOutput: jest.fn().mockReturnValue({
    setTitle: jest.fn().mockReturnThis(),
    setWidth: jest.fn().mockReturnThis(),
    setHeight: jest.fn().mockReturnThis()
  }),
  createTemplateFromFile: jest.fn().mockReturnValue({
    evaluate: jest.fn().mockReturnValue({
      setTitle: jest.fn().mockReturnThis(),
      setWidth: jest.fn().mockReturnThis(),
      setHeight: jest.fn().mockReturnThis()
    })
  })
};

global.UrlFetchApp = {
  fetch: jest.fn().mockReturnValue({
    getContentText: jest.fn().mockReturnValue('{}'),
    getResponseCode: jest.fn().mockReturnValue(200)
  })
};

global.DriveApp = {
  getFileById: jest.fn(),
  getFolderById: jest.fn(),
  getFiles: jest.fn().mockReturnValue({
    hasNext: jest.fn().mockReturnValue(false),
    next: jest.fn()
  }),
  getFolders: jest.fn().mockReturnValue({
    hasNext: jest.fn().mockReturnValue(false),
    next: jest.fn()
  })
};

global.Utilities = {
  formatDate: jest.fn().mockReturnValue('01/01/2025'),
  formatString: jest.fn((...args) => args.slice(1).reduce((str, arg, i) => 
    str.replace(new RegExp(`\\{${i}\\}`, 'g'), arg), args[0])),
  sleep: jest.fn(),
  randomString: jest.fn().mockReturnValue('random-string'),
  base64Encode: jest.fn(str => Buffer.from(str).toString('base64')),
  base64Decode: jest.fn(str => Buffer.from(str, 'base64').toString()),
  DigestAlgorithm: {
    MD5: 'MD5',
    SHA_1: 'SHA_1',
    SHA_256: 'SHA_256',
    SHA_512: 'SHA_512'
  }
};

global.GmailApp = {
  sendEmail: jest.fn(),
  createDraft: jest.fn()
};

// Mock console for testing
global.console = {
  ...console,
  log: jest.fn(),
  info: jest.fn(),
  warn: jest.fn(),
  error: jest.fn()
};