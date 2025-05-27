# Local Testing Setup for YSLv6Hub

This project now includes a local testing environment using Jest and TypeScript, allowing you to test your Google Apps Script code without deploying to Google.

## Setup Instructions

1. **Install Dependencies**
   ```bash
   npm install
   ```

2. **Run Tests**
   ```bash
   npm test
   ```

3. **Run Tests in Watch Mode**
   ```bash
   npm run test:watch
   ```

4. **Run Tests with Coverage**
   ```bash
   npm run test:coverage
   ```

## Project Structure

```
yslv6hub-gs/
├── tests/
│   ├── setup.ts              # Jest setup with GAS API mocks
│   ├── ErrorHandling.test.ts # Example test for ErrorHandling module
│   └── MenuSystem.test.ts    # Example test for MenuSystem module
├── package.json              # Node dependencies and scripts
├── tsconfig.json            # TypeScript configuration
└── jest.config.js           # Jest configuration (in package.json)
```

## Writing Tests

### Example Test Structure

```typescript
import '../YourModule';

describe('YourModule', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  it('should do something', () => {
    // Arrange
    const mockSheet = {
      getName: jest.fn(() => 'TestSheet'),
      getRange: jest.fn(),
    };
    
    const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    (mockSpreadsheet.getSheetByName as jest.Mock).mockReturnValue(mockSheet);
    
    // Act
    yourFunction();
    
    // Assert
    expect(mockSheet.getRange).toHaveBeenCalled();
  });
});
```

## Available Mocks

The `tests/setup.ts` file provides mocks for common Google Apps Script APIs:

- **SpreadsheetApp**: Full mock including getActiveSpreadsheet, getUi, create
- **HtmlService**: For testing HTML output and dialogs
- **PropertiesService**: For testing script and user properties
- **DriveApp**: For testing file and folder operations
- **MailApp**: For testing email functionality
- **Utilities**: For testing date formatting and string utilities
- **Logger**: For testing logging functionality

## Running Tests Before Deployment

Add this to your workflow:

1. Make code changes
2. Run `npm test` to verify functionality
3. Run `npm run lint` to check code style
4. Run `npm run typecheck` to verify TypeScript types
5. Deploy with `clasp push`

## Benefits

- **Fast Feedback**: Tests run in seconds, not minutes
- **No Deployment Required**: Test locally without pushing to Google
- **Debugging**: Use debugger and breakpoints in your IDE
- **Coverage Reports**: See which code paths are tested
- **Regression Prevention**: Catch bugs before deployment

## Tips

1. **Mock External Dependencies**: Always mock Google Apps Script APIs
2. **Test Business Logic**: Focus on testing your custom logic
3. **Use TypeScript**: Leverage type safety in tests
4. **Keep Tests Simple**: Each test should verify one behavior
5. **Run Tests Often**: Use watch mode during development