/**
 * Google Sheets Interaction Utilities for E2E Tests
 */

import { Page, ElementHandle } from 'puppeteer';
import { wait } from './wait';

export class SheetsHelper {
  constructor(private page: Page) {}

  /**
   * Open a specific spreadsheet by ID
   */
  async openSpreadsheet(spreadsheetId: string): Promise<void> {
    const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`;
    await this.page.goto(url, { waitUntil: 'networkidle2' });
    
    // Wait for the spreadsheet to load
    await this.page.waitForSelector('#docs-editor-container', { timeout: 30000 });
    await wait(2000); // Additional wait for full load
  }

  /**
   * Create a new spreadsheet
   */
  async createNewSpreadsheet(): Promise<string> {
    await this.page.goto('https://sheets.google.com/create', { waitUntil: 'networkidle2' });
    await this.page.waitForSelector('#docs-editor-container', { timeout: 30000 });
    
    // Get the spreadsheet ID from the URL
    const url = this.page.url();
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    return match ? match[1] : '';
  }

  /**
   * Open the YSL Hub menu
   */
  async openYSLMenu(): Promise<void> {
    // Click on Extensions menu
    await this.clickMenuItem('Extensions');
    await wait(1000);
    
    // Look for YSL Hub Enhanced menu item
    await this.clickMenuItem('YSL Hub Enhanced');
    await wait(500);
  }

  /**
   * Click a menu item by text
   */
  async clickMenuItem(text: string): Promise<void> {
    // Google Sheets uses custom menu elements
    const menuSelector = `[role="menuitem"]:has-text("${text}"), [role="menuitemcheckbox"]:has-text("${text}"), [role="menuitemradio"]:has-text("${text}")`;
    
    // First try to find in visible menus
    let element = await this.page.$(menuSelector);
    
    if (!element) {
      // If not found, it might be in the menu bar
      const menuBarSelector = `[role="menubar"] [role="menuitem"]:has-text("${text}")`;
      element = await this.page.$(menuBarSelector);
    }
    
    if (element) {
      await element.click();
    } else {
      throw new Error(`Menu item "${text}" not found`);
    }
  }

  /**
   * Get cell value
   */
  async getCellValue(cell: string): Promise<string> {
    // Click on the cell
    await this.selectCell(cell);
    
    // Get value from formula bar
    const formulaBar = await this.page.$('#t-formula-bar-input-container input');
    if (formulaBar) {
      return await this.page.evaluate(el => el.value, formulaBar);
    }
    return '';
  }

  /**
   * Set cell value
   */
  async setCellValue(cell: string, value: string): Promise<void> {
    await this.selectCell(cell);
    await this.page.keyboard.type(value);
    await this.page.keyboard.press('Enter');
    await wait(500); // Wait for value to be saved
  }

  /**
   * Select a cell (e.g., "A1", "B2")
   */
  async selectCell(cell: string): Promise<void> {
    // Click on name box
    const nameBox = await this.page.$('#t-name-box-input');
    if (nameBox) {
      await nameBox.click();
      
      // Clear and type cell reference
      await this.page.keyboard.down('Control');
      await this.page.keyboard.press('A');
      await this.page.keyboard.up('Control');
      await this.page.keyboard.type(cell);
      await this.page.keyboard.press('Enter');
      
      await wait(300);
    }
  }

  /**
   * Select a range (e.g., "A1:B10")
   */
  async selectRange(range: string): Promise<void> {
    await this.selectCell(range);
  }

  /**
   * Get sheet names
   */
  async getSheetNames(): Promise<string[]> {
    const sheetTabs = await this.page.$$('.docs-sheet-tab-name');
    const names: string[] = [];
    
    for (const tab of sheetTabs) {
      const name = await this.page.evaluate(el => el.textContent || '', tab);
      names.push(name.trim());
    }
    
    return names;
  }

  /**
   * Switch to a specific sheet
   */
  async switchToSheet(sheetName: string): Promise<void> {
    const sheetTab = await this.page.$(`[role="tab"]:has-text("${sheetName}")`);
    if (sheetTab) {
      await sheetTab.click();
      await wait(500);
    } else {
      throw new Error(`Sheet "${sheetName}" not found`);
    }
  }

  /**
   * Wait for and dismiss any dialogs
   */
  async dismissDialog(): Promise<void> {
    try {
      // Look for common dialog close buttons
      const closeButton = await this.page.$('[aria-label="Close"], [aria-label="Cancel"], button:has-text("Close"), button:has-text("Cancel")');
      if (closeButton) {
        await closeButton.click();
        await wait(500);
      }
    } catch (error) {
      // Ignore if no dialog found
    }
  }

  /**
   * Take a screenshot with a descriptive name
   */
  async screenshot(name: string): Promise<void> {
    const screenshotDir = process.env.SCREENSHOT_DIR || './e2e/screenshots';
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const filename = `${name}-${timestamp}.png`;
    await this.page.screenshot({ 
      path: `${screenshotDir}/${filename}`, 
      fullPage: true 
    });
  }

  /**
   * Press a keyboard key
   */
  async pressKey(key: any): Promise<void> {
    await this.page.keyboard.press(key);
  }
}