/**
 * Wait utility for Puppeteer tests
 */

export async function wait(ms: number): Promise<void> {
  return new Promise(resolve => setTimeout(resolve, ms));
}