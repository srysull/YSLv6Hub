/**
 * Google Authentication Helper for E2E Tests
 */

import { Page } from 'puppeteer';

export class GoogleAuth {
  private email: string;
  private password: string;

  constructor() {
    this.email = process.env.GOOGLE_EMAIL!;
    this.password = process.env.GOOGLE_PASSWORD!;
  }

  /**
   * Sign in to Google account
   */
  async signIn(page: Page): Promise<void> {
    console.log('Signing in to Google...');
    
    // Go to Google sign-in page
    await page.goto('https://accounts.google.com/signin');
    
    // Enter email
    await page.waitForSelector('input[type="email"]');
    await page.type('input[type="email"]', this.email);
    await page.click('#identifierNext');
    
    // Wait for password field and enter password
    await page.waitForSelector('input[type="password"]', { visible: true });
    await new Promise(resolve => setTimeout(resolve, 1000)); // Small delay for stability
    await page.type('input[type="password"]', this.password);
    await page.click('#passwordNext');
    
    // Wait for navigation to complete
    await page.waitForNavigation({ waitUntil: 'networkidle2' });
    
    console.log('Successfully signed in to Google');
  }

  /**
   * Check if user is already signed in
   */
  async isSignedIn(page: Page): Promise<boolean> {
    try {
      // Navigate to Google Sheets and check if we're redirected to sign-in
      await page.goto('https://sheets.google.com', { waitUntil: 'networkidle2' });
      const url = page.url();
      return !url.includes('accounts.google.com');
    } catch (error) {
      return false;
    }
  }

  /**
   * Sign in if not already signed in
   */
  async ensureSignedIn(page: Page): Promise<void> {
    const signedIn = await this.isSignedIn(page);
    if (!signedIn) {
      await this.signIn(page);
    }
  }
}