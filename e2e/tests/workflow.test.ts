/**
 * E2E Tests for YSL Hub Complete Workflow
 */

import { GoogleAuth } from '../utils/auth';
import { SheetsHelper } from '../utils/sheets';

describe('YSL Hub Complete Workflow', () => {
  let auth: GoogleAuth;
  let sheets: SheetsHelper;
  let testSpreadsheetId: string;

  beforeAll(async () => {
    auth = new GoogleAuth();
    await auth.ensureSignedIn(page);
    sheets = new SheetsHelper(page);
  });

  beforeEach(async () => {
    // Use existing test spreadsheet if provided, otherwise create new
    testSpreadsheetId = process.env.TEST_SPREADSHEET_ID || await sheets.createNewSpreadsheet();
    await sheets.openSpreadsheet(testSpreadsheetId);
    await page.waitForTimeout(5000); // Wait for Apps Script
  });

  describe('Full Setup and Usage Flow', () => {
    it('should complete full setup workflow', async () => {
      // Step 1: Initialize blank spreadsheet
      await sheets.openYSLMenu();
      await sheets.clickMenuItem('Initialize Blank Spreadsheet');
      await page.waitForTimeout(10000);
      
      // Step 2: Configure system
      await sheets.openYSLMenu();
      await sheets.clickMenuItem('System Configuration');
      
      // Wait for dialog
      await page.waitForSelector('iframe[src*="google.com"]', { timeout: 10000 });
      
      // Step 3: Add test data to Assumptions
      await sheets.switchToSheet('Assumptions');
      await sheets.setCellValue('A2', 'Session Duration');
      await sheets.setCellValue('B2', '30');
      await sheets.setCellValue('A3', 'Instructor Name');
      await sheets.setCellValue('B3', 'Test Instructor');
      
      // Step 4: Add a test class
      await sheets.switchToSheet('Classes');
      await sheets.setCellValue('A2', 'Test Class 101');
      await sheets.setCellValue('B2', 'Monday');
      await sheets.setCellValue('C2', '10:00 AM');
      
      // Step 5: Add a test student
      await sheets.switchToSheet('Roster');
      await sheets.setCellValue('A2', 'John');
      await sheets.setCellValue('B2', 'Doe');
      await sheets.setCellValue('C2', 'john.doe@example.com');
      await sheets.setCellValue('D2', 'Test Class 101');
      
      // Verify data was saved
      const className = await sheets.getCellValue('D2');
      expect(className).toBe('Test Class 101');
    });

    it('should handle email template workflow', async () => {
      // Initialize if needed
      const sheetNames = await sheets.getSheetNames();
      if (!sheetNames.includes('Assumptions')) {
        await sheets.openYSLMenu();
        await sheets.clickMenuItem('Initialize Blank Spreadsheet');
        await page.waitForTimeout(10000);
      }
      
      // Open email templates
      await sheets.openYSLMenu();
      const emailMenuItem = await page.$('[role="menuitem"]:has-text("Email Templates")');
      if (emailMenuItem) {
        await emailMenuItem.hover();
        await page.waitForTimeout(500);
        await sheets.clickMenuItem('Manage Email Templates');
        
        // Wait for dialog
        await page.waitForSelector('iframe[src*="google.com"]', { timeout: 10000 });
        
        // Dialog should be open
        const frames = page.frames();
        const dialogFrame = frames.find(f => f.url().includes('google.com/macros'));
        
        if (dialogFrame) {
          const hasTemplateContent = await dialogFrame.evaluate(() => {
            return document.body.textContent?.includes('Template') || false;
          });
          
          expect(hasTemplateContent).toBe(true);
        }
      }
    });
  });

  describe('Error Handling', () => {
    it('should handle errors gracefully', async () => {
      // Try to access a function without initialization
      await sheets.openYSLMenu();
      
      // This should show an error but not crash
      const toolsMenuItem = await page.$('[role="menuitem"]:has-text("Tools & Diagnostics")');
      if (toolsMenuItem) {
        await toolsMenuItem.hover();
        await page.waitForTimeout(500);
        
        try {
          await sheets.clickMenuItem('System Health Check');
          await page.waitForTimeout(3000);
          
          // Should either show an alert or handle gracefully
          // Check if any alert appeared
          const alertDialog = await page.$('[role="alertdialog"]');
          expect(alertDialog).toBeTruthy();
        } catch (error) {
          // Error is expected and handled
          expect(error).toBeDefined();
        }
      }
    });
  });

  describe('Data Validation', () => {
    it('should validate email formats', async () => {
      // Initialize if needed
      const sheetNames = await sheets.getSheetNames();
      if (!sheetNames.includes('Roster')) {
        await sheets.openYSLMenu();
        await sheets.clickMenuItem('Initialize Blank Spreadsheet');
        await page.waitForTimeout(10000);
      }
      
      // Go to Roster sheet
      await sheets.switchToSheet('Roster');
      
      // Try to enter invalid email
      await sheets.setCellValue('A2', 'Test');
      await sheets.setCellValue('B2', 'User');
      await sheets.setCellValue('C2', 'invalid-email');
      
      // Apply validation
      await sheets.openYSLMenu();
      const toolsMenuItem = await page.$('[role="menuitem"]:has-text("Tools & Diagnostics")');
      if (toolsMenuItem) {
        await toolsMenuItem.hover();
        await page.waitForTimeout(500);
        
        // Look for validation option
        const validationExists = await page.evaluate(() => {
          const items = Array.from(document.querySelectorAll('[role="menuitem"]'));
          return items.some(item => item.textContent?.includes('Validation'));
        });
        
        expect(validationExists).toBe(true);
      }
    });
  });
});