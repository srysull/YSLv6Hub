/**
 * E2E Tests for YSL Hub Menu System
 */

import { GoogleAuth } from '../utils/auth';
import { SheetsHelper } from '../utils/sheets';

describe('YSL Hub Menu System', () => {
  let auth: GoogleAuth;
  let sheets: SheetsHelper;
  let testSpreadsheetId: string;

  beforeAll(async () => {
    auth = new GoogleAuth();
    await auth.ensureSignedIn(page);
    sheets = new SheetsHelper(page);
  });

  beforeEach(async () => {
    // Create a new test spreadsheet for each test
    testSpreadsheetId = await sheets.createNewSpreadsheet();
    console.log(`Created test spreadsheet: ${testSpreadsheetId}`);
    
    // Wait for Apps Script to load
    await page.waitForTimeout(5000);
  });

  afterEach(async () => {
    // Take screenshot if test failed
    if (jasmine.currentTest.failedExpectations.length > 0) {
      await sheets.screenshot(`failed-${jasmine.currentTest.fullName}`);
    }
  });

  describe('Menu Creation', () => {
    it('should create YSL Hub Enhanced menu', async () => {
      // The menu should be created automatically on open
      await sheets.clickMenuItem('Extensions');
      
      // Check if YSL Hub Enhanced menu exists
      const menuExists = await page.evaluate(() => {
        const menuItems = Array.from(document.querySelectorAll('[role="menuitem"]'));
        return menuItems.some(item => item.textContent?.includes('YSL Hub Enhanced'));
      });
      
      expect(menuExists).toBe(true);
    });

    it('should have all main menu items', async () => {
      await sheets.openYSLMenu();
      
      // Check for main menu items
      const expectedItems = [
        'System Configuration',
        'Initialize Blank Spreadsheet',
        'Email Templates',
        'Tools & Diagnostics'
      ];
      
      for (const item of expectedItems) {
        const itemExists = await page.evaluate((text) => {
          const menuItems = Array.from(document.querySelectorAll('[role="menuitem"]'));
          return menuItems.some(el => el.textContent?.includes(text));
        }, item);
        
        expect(itemExists).toBe(true);
      }
    });
  });

  describe('Initialize Blank Spreadsheet', () => {
    it('should create all required sheets', async () => {
      await sheets.openYSLMenu();
      await sheets.clickMenuItem('Initialize Blank Spreadsheet');
      
      // Wait for initialization to complete
      await page.waitForTimeout(10000);
      
      // Check if all required sheets were created
      const sheetNames = await sheets.getSheetNames();
      const expectedSheets = ['Assumptions', 'Classes', 'Roster', 'SystemLogs'];
      
      for (const sheetName of expectedSheets) {
        expect(sheetNames).toContain(sheetName);
      }
    });

    it('should set up proper formatting in Assumptions sheet', async () => {
      await sheets.openYSLMenu();
      await sheets.clickMenuItem('Initialize Blank Spreadsheet');
      
      // Wait for initialization
      await page.waitForTimeout(10000);
      
      // Switch to Assumptions sheet
      await sheets.switchToSheet('Assumptions');
      
      // Check if headers are present
      const a1Value = await sheets.getCellValue('A1');
      expect(a1Value).toBe('Assumption');
      
      const b1Value = await sheets.getCellValue('B1');
      expect(b1Value).toBe('Value');
    });
  });

  describe('System Configuration Dialog', () => {
    it('should open configuration dialog', async () => {
      await sheets.openYSLMenu();
      await sheets.clickMenuItem('System Configuration');
      
      // Wait for dialog to appear
      await page.waitForSelector('iframe[src*="google.com"]', { timeout: 10000 });
      
      // Switch to iframe context
      const frames = page.frames();
      const dialogFrame = frames.find(f => f.url().includes('google.com/macros'));
      
      if (dialogFrame) {
        // Check if dialog content loaded
        const titleExists = await dialogFrame.evaluate(() => {
          return document.body.textContent?.includes('Configuration') || false;
        });
        
        expect(titleExists).toBe(true);
      }
    });
  });

  describe('Email Templates', () => {
    it('should open email templates submenu', async () => {
      await sheets.openYSLMenu();
      
      // Hover over Email Templates to open submenu
      const emailMenuItem = await page.$('[role="menuitem"]:has-text("Email Templates")');
      if (emailMenuItem) {
        await emailMenuItem.hover();
        await page.waitForTimeout(500);
        
        // Check for submenu items
        const manageTemplatesExists = await page.evaluate(() => {
          const items = Array.from(document.querySelectorAll('[role="menuitem"]'));
          return items.some(item => item.textContent?.includes('Manage Email Templates'));
        });
        
        expect(manageTemplatesExists).toBe(true);
      }
    });
  });
});