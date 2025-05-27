/**
 * E2E Test Setup
 * Configures Puppeteer and handles common setup tasks
 */

import * as dotenv from 'dotenv';
import * as fs from 'fs';
import * as path from 'path';

// Load environment variables
dotenv.config();

// Verify required environment variables
const requiredEnvVars = ['GOOGLE_EMAIL', 'GOOGLE_PASSWORD'];
const missingVars = requiredEnvVars.filter(varName => !process.env[varName]);

if (missingVars.length > 0) {
  console.error('Missing required environment variables:', missingVars);
  console.error('Please copy .env.example to .env and fill in your test credentials');
  process.exit(1);
}

// Create screenshot directory if it doesn't exist
const screenshotDir = process.env.SCREENSHOT_DIR || './e2e/screenshots';
if (!fs.existsSync(screenshotDir)) {
  fs.mkdirSync(screenshotDir, { recursive: true });
}

// Global test timeout
jest.setTimeout(parseInt(process.env.E2E_TIMEOUT || '30000'));

// Add custom matchers if needed
expect.extend({
  async toHaveMenuItemWithText(received: any, text: string) {
    try {
      await received.waitForSelector(`[role="menuitem"]:has-text("${text}")`, { timeout: 5000 });
      return {
        pass: true,
        message: () => `Found menu item with text "${text}"`,
      };
    } catch (error) {
      return {
        pass: false,
        message: () => `Could not find menu item with text "${text}"`,
      };
    }
  },
});

// Global error handling
process.on('unhandledRejection', (error: Error) => {
  console.error('Unhandled rejection:', error);
});

// Take screenshot on test failure
global.afterEach(async () => {
  if (global.page && global.jasmine.currentTest.failedExpectations.length > 0) {
    const testName = global.jasmine.currentTest.fullName.replace(/\s+/g, '_');
    const screenshotPath = path.join(screenshotDir, `${testName}-${Date.now()}.png`);
    await global.page.screenshot({ path: screenshotPath, fullPage: true });
    console.log(`Screenshot saved: ${screenshotPath}`);
  }
});

// Close browser after all tests
global.afterAll(async () => {
  if (global.browser) {
    await global.browser.close();
  }
});