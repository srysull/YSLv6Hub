# YSL Hub E2E Testing with Puppeteer

This directory contains end-to-end tests for the YSL Hub Google Apps Script project using Puppeteer and Jest.

## Setup

1. **Install dependencies**:
   ```bash
   npm install
   ```

2. **Configure environment**:
   ```bash
   cp .env.example .env
   ```
   Edit `.env` and add your Google test account credentials:
   - `GOOGLE_EMAIL`: Your test Google account email
   - `GOOGLE_PASSWORD`: Your test account password (use app-specific password if 2FA is enabled)
   - `TEST_SPREADSHEET_ID`: (Optional) ID of a test spreadsheet to reuse

3. **Create app-specific password** (if using 2FA):
   - Go to https://myaccount.google.com/security
   - Enable 2-Step Verification
   - Generate an app-specific password
   - Use this password in your `.env` file

## Running Tests

### Run all E2E tests:
```bash
npm run test:e2e
```

### Run tests with visible browser (headful mode):
```bash
npm run test:e2e:headful
```

### Run tests in debug mode (slow, visible):
```bash
npm run test:e2e:debug
```

### Run specific test file:
```bash
npm run test:e2e -- menu.test.ts
```

## Test Structure

```
e2e/
├── tests/           # Test files
│   ├── menu.test.ts     # Menu system tests
│   └── workflow.test.ts # Complete workflow tests
├── utils/           # Helper utilities
│   ├── auth.ts          # Google authentication
│   ├── sheets.ts        # Sheets interaction helpers
│   └── testData.ts      # Test data management
└── setup.ts         # Jest setup and configuration
```

## Writing Tests

### Basic test structure:
```typescript
import { GoogleAuth } from '../utils/auth';
import { SheetsHelper } from '../utils/sheets';

describe('My Feature', () => {
  let sheets: SheetsHelper;

  beforeAll(async () => {
    const auth = new GoogleAuth();
    await auth.ensureSignedIn(page);
    sheets = new SheetsHelper(page);
  });

  it('should do something', async () => {
    await sheets.openSpreadsheet('your-spreadsheet-id');
    await sheets.setCellValue('A1', 'Hello');
    const value = await sheets.getCellValue('A1');
    expect(value).toBe('Hello');
  });
});
```

### Using test data helpers:
```typescript
import { TestDataHelper } from '../utils/testData';

const testData = new TestDataHelper(sheets);
const { classes, students } = TestDataHelper.generateSampleData();

await testData.setupAssumptions();
await testData.addTestClasses(classes);
await testData.addTestStudents(students);
```

## Helper Methods

### SheetsHelper
- `openSpreadsheet(id)` - Open a spreadsheet by ID
- `createNewSpreadsheet()` - Create a new blank spreadsheet
- `getCellValue(cell)` - Get value from a cell (e.g., 'A1')
- `setCellValue(cell, value)` - Set cell value
- `selectCell(cell)` - Select a specific cell
- `selectRange(range)` - Select a range (e.g., 'A1:B10')
- `getSheetNames()` - Get all sheet tab names
- `switchToSheet(name)` - Switch to a specific sheet
- `openYSLMenu()` - Open the YSL Hub menu
- `clickMenuItem(text)` - Click a menu item by text
- `screenshot(name)` - Take a screenshot

### TestDataHelper
- `setupAssumptions()` - Set up default assumptions
- `addTestClasses(classes)` - Add test classes
- `addTestStudents(students)` - Add test students
- `clearAllData()` - Clear all test data
- `verifyDataIntegrity()` - Check if data was saved correctly

## Best Practices

1. **Use test accounts**: Never use production Google accounts
2. **Clean up**: Clear test data after tests
3. **Wait for operations**: Google Sheets operations are async
4. **Take screenshots**: Capture state on failures
5. **Handle timeouts**: Network operations may be slow

## Troubleshooting

### Authentication Issues
- Ensure app-specific password is used with 2FA
- Check that less secure app access is enabled (if not using 2FA)
- Verify credentials in `.env` file

### Timeout Issues
- Increase `E2E_TIMEOUT` in `.env`
- Add explicit waits after operations
- Use `networkidle2` for page loads

### Element Not Found
- Google Sheets UI may change
- Use role-based selectors when possible
- Add waits before interacting with elements

## Screenshots

Failed tests automatically save screenshots to `./e2e/screenshots/` with descriptive names including test name and timestamp.