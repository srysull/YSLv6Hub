# YSLv6Hub Enhanced

YSLv6Hub is a comprehensive Google Workspace system for managing the Youth Swim Lessons (YSL) program at PenBay YMCA. This enhanced version adds several key improvements to make the system more robust, user-friendly, and feature-rich.

## Project Structure

This project uses clasp's direct TypeScript compilation capability to deploy TypeScript files directly to Google Apps Script.

### Directory Organization
```
/Users/galagrove/Projects/YSL/
├── YSLv6Hub-main/          # Main development repository (this folder)
│   ├── *.ts                # TypeScript source files
│   ├── *.gs                # Compiled Google Apps Script files
│   ├── tests/              # Unit tests with Jest
│   ├── e2e/                # End-to-end tests with Puppeteer
│   └── templates/          # Excel templates
├── archive/                # Old versions and backups
└── documentation/          # Business analysis and documentation
```

### Key Files
- `00_MenuSystem.ts` - Centralized menu implementation
- `00_TriggerFunctions.ts` - Entry points for menu and edit triggers
- `01_Globals.ts` - Common functions and utilities
- `02_ErrorHandling.ts` - Error handling and logging
- `03_VersionControl.ts` - Version control and diagnostics
- `04_AdministrativeModule.ts` - System administration
- `05_MenuWrappers.ts` - Menu function wrappers
- `15_DynamicInstructorSheet.ts` - Group Lesson Tracker generation
- `18_SyncFunctions.ts` - Data synchronization functions

## Key Features

### Core Functionality
- Class management and instructor resources
- Student roster management and assessment tracking
- Email communications with parents and participants
- Reporting on student progress
- System administration and configuration

### New Enhancements
- **Fixed Sync Data Function**: Properly handles syncing data between Group Lesson Tracker and SwimmerSkills sheets
- **Centralized Menu System**: Prevents duplicate menu issues and provides consistent menu access
- **Blank Spreadsheet Initializer**: Creates all required sheets and structure from a completely blank spreadsheet
- **Email Templates System**: Reusable email templates with placeholders for personalization
- **Input Validation**: Enhanced data validation throughout the system for improved data integrity
- **Local Testing Framework**: Jest unit tests and Puppeteer E2E tests for development

## Development Workflow

### Setup
1. Clone or use the repository at `/Users/galagrove/Projects/YSL/YSLv6Hub-main`
2. Install dependencies: `npm install`
3. Copy `.env.example` to `.env` for E2E testing credentials

### Development
1. Edit the TypeScript files in this directory
2. Run unit tests: `npm test`
3. Run E2E tests: `npm run test:e2e`
4. Deploy to Google Apps Script: `clasp push`
5. Access your project in the Google Apps Script editor

### Testing
- **Unit Tests**: `npm test` - Fast local testing with mocked GAS APIs
- **E2E Tests**: `npm run test:e2e` - Full browser testing with Puppeteer
- **Watch Mode**: `npm run test:watch` - Auto-run tests on file changes
- **Coverage**: `npm run test:coverage` - Generate test coverage reports

## Getting Started

### For New Spreadsheets
1. Open a new or blank Google Sheet
2. Go to Extensions > Apps Script
3. Copy all TypeScript files from this repository into the Apps Script editor
4. Save and close the editor
5. Refresh your spreadsheet
6. Use the "YSL Hub Enhanced" menu > "Initialize Blank Spreadsheet"
7. Follow the initialization wizard

### For Existing YSL Hub Spreadsheets
1. Go to Extensions > Apps Script
2. Copy all TypeScript files from this repository into the Apps Script editor
3. Save and close the editor
4. Refresh your spreadsheet
5. Use the "YSL Hub Enhanced" menu > "System Configuration"

## Module Structure

- `00_MenuSystem.ts`: Centralized menu creation
- `01_Globals.ts`: Global functions and utilities
- `02_ErrorHandling.ts`: Centralized error handling and logging
- `03_VersionControl.ts`: Version management and updates
- `04_AdministrativeModule.ts`: System initialization and configuration
- `05_MenuWrappers.ts`: Menu creation and event handlers
- `06_DataIntegrationModule.ts`: Data processing and management
- `07_CommunicationModule.ts`: Email and notifications
- `08_ReportingModule.ts`: Assessment reports generation
- `09_UserGuideModule.ts`: User guide and documentation
- `10_HistoryModule.ts`: History tracking and reporting
- `11_SessionTransitionModule.ts`: Session transition management
- `12_InstructorResourceModule.ts`: Instructor-specific tools
- `13_VersionControlActions.ts`: Version control actions
- `14_DebugModule.ts`: Debugging utilities
- `15_DynamicInstructorSheet.ts`: Group Lesson Tracker generation
- `16_InstallTrigger.ts`: Installation triggers
- `17_MenuFix.ts`: Menu system fixes
- `18_SyncFunctions.ts`: Data synchronization

## Deployment

This project is configured to deploy to Google Apps Script ID: `17mxN2QUfg6sWx7X88TYeJ_ceaxjp8g07b6MivqFzqnv0-u8Y60tEM9FV`

To deploy:
```bash
clasp push
```

## Support

For support or to report issues, please contact:
- ssullivan@penbayymca.org
- GitHub: https://github.com/srysull/YSLv6Hub

## Version

YSLv6Hub v6.0.5
Last Updated: May 27, 2025