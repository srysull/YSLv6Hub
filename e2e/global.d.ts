import { Page, Browser } from 'puppeteer';

declare global {
  var page: Page;
  var browser: Browser;
  
  namespace jasmine {
    interface CurrentTest {
      fullName: string;
      failedExpectations: any[];
    }
    
    var currentTest: CurrentTest;
  }
}

export {};