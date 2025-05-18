# YSLv6Hub

> **Note:** For Claude Code assistance when resuming work, refer to both this main project folder AND the backup documentation in the Google Drive folder. See "Project Documentation" section below.

## Project Overview

YSLv6Hub is a refactored TypeScript implementation of the YMCA Swim Lessons application, providing improved organization, type safety, and modern development practices while preserving critical functionality. It streamlines the management of swim lessons with features for registration data handling, class tracking, skills assessment, attendance logging, communications, and reporting.

## Critical Functionality

The most crucial functionality that must be preserved is the **GroupsTracker and SwimmerSkills integration**, specifically:

1. GroupsTracker sheet generation with formatted templates
2. Student data population from roster information
3. Skills population based on class level/stage
4. Bidirectional synchronization between GroupsTracker and SwimmerSkills sheets

Detailed documentation of this functionality is available in the backup folder (see below).

## Key Features

- Smart registration data import with date-range preservation
- On-demand data loading for class tracking
- Bidirectional synchronization between tracking sheets
- HTML-formatted communications with attachment support
- Comprehensive dashboard for system monitoring
- Role-based access control
- Session transition with student continuity

## Repository Information

- **Repository URL:** [Private GitHub Repository]
- **Branch Strategy:**
  - `main`: Production-ready code
  - `develop`: Integration branch for feature development
  - `feature/XX-feature-name`: Individual feature branches

### Getting Started with the Repository

```bash
# Clone the repository
git clone [repository-url] yslv6hub
cd yslv6hub

# Install dependencies
npm install

# Configure CLASP with the Google Apps Script project
clasp login
clasp clone [script-id] --rootDir ./dist
```

## Project Structure

```
/Users/galagrove/yslv6hub/
├── src/                    # Source code
│   ├── 00_System.ts        # Main entry point, initialization, menu creation
│   ├── 01_Core.ts          # Core utilities, error handling, caching, events
│   ├── 02_DataAccess.ts    # Sheet access, data retrieval, manipulation
│   ├── 03_Templates.ts     # UI templates and components
│   ├── 04_GroupsTracker.ts # GroupsTracker sheet creation and management
│   ├── 05_SkillsSync.ts    # Bidirectional sync with SwimmerSkills
│   └── [Additional modules]
├── test/                   # Jest test files
│   ├── mocks/              # Mock implementations of Google Apps Script services
│   ├── 00_System.test.ts   # Tests for System module
│   ├── 01_Core.test.ts     # Tests for Core module
│   └── [Additional tests]
├── dist/                   # Compiled JavaScript output (for CLASP)
├── types/                  # Custom type definitions
│   └── google-apps-script.d.ts # Google Apps Script type definitions
├── .eslintrc.js            # ESLint configuration
├── .eslintignore           # Files to ignore in linting
├── tsconfig.json           # TypeScript configuration
├── package.json            # Project dependencies and scripts
├── jest.config.js          # Jest testing configuration
├── .clasp.json             # CLASP configuration
├── .claspignore            # Files to ignore in CLASP uploads
├── .gitignore              # Files to ignore in Git
├── PROJECT_TASKS.md        # Detailed task list and progress
└── README.md               # This file
```

## Project Documentation

Documentation exists in two locations:

1. **Main Project Folder (`/Users/galagrove/yslv6hub/`):** 
   Contains implementation code, PROJECT_TASKS.md tracking progress, and active development files.

2. **Backup Documentation (`/Users/galagrove/Library/CloudStorage/GoogleDrive-ssullivan@penbayymca.org/My Drive/SRS YSLv6Hub/backup_for_claude/`):**
   Contains detailed analysis and skeleton implementations, including:
   - `CRITICAL_FUNCTIONALITY_DOCUMENTATION.md`: Details of the critical GroupsTracker/SwimmerSkills integration
   - `MIGRATION_PLAN.md`: Comprehensive plan for migrating the functionality
   - `04_GroupsTracker.ts`: Skeleton implementation of GroupsTracker functionality
   - `05_SkillsSync.ts`: Skeleton implementation of bidirectional sync
   - `PROJECT_STATUS.md`: Current project status and next steps
   - Original code backups for reference

**IMPORTANT:** When resuming work with Claude Code, point to this main project folder, but also mention the backup documentation location to ensure Claude has full context.

## Sheet Structure

- **YSLv6Hub [SHEET]**: Main dashboard and implementation guide
- **RegistrationInfo [SHEET]**: Registration data
- **GroupsTracker [SHEET]**: Class tracking sheets
- **SwimmerSkills [SHEET]**: Skills assessment tracking
- **SwimmerLog [SHEET]**: Attendance and activity logging
- **CommsHub [SHEET]**: Communications management
- **SystemLog [SHEET]**: System activity tracking

## Development Environment Setup

### Prerequisites

1. **Node.js and npm**: Version 14.x or later recommended
   ```bash
   # Check versions
   node -v
   npm -v
   ```

2. **Git**: For version control
   ```bash
   # Check version
   git --version
   ```

3. **Google Account**: With access to Google Apps Script and necessary permissions

### Initial Setup

1. **Install global tools**:
   ```bash
   # Install CLASP globally
   npm install -g @google/clasp
   
   # Login to Google
   clasp login
   ```

2. **Project setup**:
   ```bash
   # Install project dependencies
   npm install
   
   # Initialize Git repository (if not cloned)
   git init
   git add .
   git commit -m "Initial commit"
   ```

3. **CLASP configuration**:
   ```bash
   # Create a new Google Apps Script project (if needed)
   clasp create --title "YSLv6Hub" --rootDir ./dist
   
   # Or link to existing project
   clasp clone [script-id] --rootDir ./dist
   ```

4. **Environment verification**:
   ```bash
   # Ensure TypeScript compiles
   npm run typecheck
   
   # Ensure ESLint runs without errors
   npm run lint
   
   # Ensure tests run
   npm test
   ```

### Development Tools Configuration

#### TypeScript Configuration (tsconfig.json)

```json
{
  "compilerOptions": {
    "target": "ES2019",
    "module": "None",
    "lib": ["ESNext"],
    "outDir": "dist",
    "rootDir": "src",
    "strict": true,
    "esModuleInterop": true,
    "skipLibCheck": true,
    "forceConsistentCasingInFileNames": true
  },
  "include": ["src/**/*.ts", "types/**/*.ts"],
  "exclude": ["node_modules", "test"]
}
```

#### ESLint Configuration (.eslintrc.js)

```javascript
module.exports = {
  root: true,
  parser: '@typescript-eslint/parser',
  plugins: ['@typescript-eslint'],
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended'
  ],
  env: {
    node: true
  },
  rules: {
    // Project-specific rules
    'no-console': 'warn',
    '@typescript-eslint/explicit-function-return-type': 'error',
    '@typescript-eslint/no-explicit-any': 'warn'
  }
};
```

#### Jest Configuration (jest.config.js)

```javascript
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  testMatch: ['**/test/**/*.test.ts'],
  collectCoverageFrom: [
    'src/**/*.ts',
    '!src/types/**/*.ts'
  ],
  setupFiles: ['./test/setup.ts'],
  moduleNameMapper: {
    // Mock Google Apps Script global objects
    '^../mocks/(.*)$': '<rootDir>/test/mocks/$1'
  },
  globals: {
    'ts-jest': {
      tsconfig: 'tsconfig.json'
    }
  }
};
```

## Development Workflow

### Daily Development Cycle

1. **Update from repository**:
   ```bash
   git pull origin develop
   npm install # If dependencies changed
   ```

2. **Create feature branch** (if starting new feature):
   ```bash
   git checkout -b feature/groupstracker-implementation
   ```

3. **Development loop**:
   ```bash
   # Make code changes in src/
   
   # Check code quality
   npm run lint
   
   # Verify type correctness
   npm run typecheck
   
   # Run tests to ensure functionality
   npm test
   
   # Build the project
   npm run build
   
   # Push to Google Apps Script to test in actual environment
   clasp push
   ```

4. **Commit changes**:
   ```bash
   git add .
   git commit -m "Implement GroupsTracker template creation"
   ```

5. **Open in Apps Script Editor** (for manual testing):
   ```bash
   clasp open
   ```

6. **Push changes** (when feature is complete):
   ```bash
   git push origin feature/groupstracker-implementation
   # Then create a pull request to develop branch
   ```

### Testing Process

#### 1. Unit Testing with Jest

Unit tests validate individual functions in isolation using mocks for Google Apps Script services.

```typescript
// Example test for GroupsTracker.createSheet()
import { GroupsTracker } from '../src/04_GroupsTracker';
import { mockSpreadsheetApp } from './mocks/SpreadsheetApp';

// Setup global mocks
global.SpreadsheetApp = mockSpreadsheetApp;

describe('GroupsTracker.createSheet', () => {
  beforeEach(() => {
    // Reset mocks before each test
    mockSpreadsheetApp.getActiveSpreadsheet.mockClear();
  });
  
  test('creates a new sheet with correct name', () => {
    // Arrange
    const mockSheet = { getName: () => 'Group Lesson Tracker' };
    const mockSS = { 
      getSheetByName: jest.fn().mockReturnValue(null),
      insertSheet: jest.fn().mockReturnValue(mockSheet)
    };
    mockSpreadsheetApp.getActiveSpreadsheet.mockReturnValue(mockSS);
    
    // Act
    const result = GroupsTracker.createSheet();
    
    // Assert
    expect(mockSS.getSheetByName).toHaveBeenCalledWith('Group Lesson Tracker');
    expect(mockSS.insertSheet).toHaveBeenCalledWith('Group Lesson Tracker');
    expect(result).toBe(mockSheet);
  });
});
```

#### 2. Integration Testing

Integration tests validate interactions between modules. These tests use broader mocks and focus on module interfaces.

```typescript
// Example integration test for SkillsSync with GroupsTracker
import { SkillsSync } from '../src/05_SkillsSync';
import { GroupsTracker } from '../src/04_GroupsTracker';
import { mockSpreadsheetApp } from './mocks/SpreadsheetApp';

global.SpreadsheetApp = mockSpreadsheetApp;

// Mock the GroupsTracker module
jest.mock('../src/04_GroupsTracker', () => ({
  GroupsTracker: {
    collectStudents: jest.fn(),
    CONFIG: { SHEET_NAME: 'Group Lesson Tracker' }
  }
}));

describe('SkillsSync integration', () => {
  test('syncData calls GroupsTracker.collectStudents', () => {
    // Arrange
    const mockSheet = { getName: () => 'Group Lesson Tracker' };
    GroupsTracker.collectStudents.mockReturnValue([]);
    
    // Act
    SkillsSync.syncData(mockSheet);
    
    // Assert
    expect(GroupsTracker.collectStudents).toHaveBeenCalledWith(mockSheet);
  });
});
```

#### 3. End-to-End Testing in Google Apps Script

After unit and integration tests pass, push changes to Google Apps Script for real-world testing:

1. Run `clasp push` to deploy to Google Apps Script
2. Use `clasp open` to open the script in browser
3. Run the relevant functions manually and verify results
4. Check functioning in actual spreadsheets

#### 4. Test Coverage Tracking

Regularly check test coverage to ensure comprehensive testing:

```bash
# Generate coverage report
npm run test:coverage
```

This creates a coverage report in the `coverage/` directory showing percentage of code covered by tests and identifying untested code.

## Critical Implementation Notes

### 1. Module Dependencies

Ensure proper module dependencies are maintained:
- `04_GroupsTracker.ts` depends on `01_Core.ts`, `02_DataAccess.ts`
- `05_SkillsSync.ts` depends on `01_Core.ts`, `02_DataAccess.ts`, `04_GroupsTracker.ts`

### 2. Google Apps Script Limitations

Be aware of these constraints when implementing:
- Max execution time: 6 minutes per execution
- Max spreadsheet operations: ~30,000 cells per minute
- No ES modules: Use namespace approach instead
- Limited modern JS features: Target ES2019 compatibility

### 3. Bidirectional Sync Behavior

The critical bidirectional sync must function exactly as before:
- GroupsTracker "End" columns → SwimmerSkills "Repeat" columns (one column to right)
- SwimmerSkills columns → GroupsTracker "Beginning" columns
- Original SwimmerSkills data must be preserved
- Color coding: X = green (completed), / = yellow (taught)

### 4. Error Handling Framework

All functions must use the standardized error handling:
```typescript
try {
  // Function logic
} catch (error) {
  Core.handleError(error, 'ModuleName.functionName', 'User-friendly error message');
  return null; // Or appropriate error indicator
}
```

### 5. UI Considerations

- Maintain familiar interface elements for user comfort
- Add loading indicators for long-running operations
- Provide clear success/failure messages
- Use consistent color coding and formatting

## npm Scripts

```json
{
  "scripts": {
    "build": "tsc",
    "watch": "tsc --watch",
    "lint": "eslint src/**/*.ts",
    "lint:fix": "eslint --fix src/**/*.ts",
    "typecheck": "tsc --noEmit",
    "test": "jest",
    "test:watch": "jest --watch",
    "test:coverage": "jest --coverage",
    "deploy": "npm run build && clasp push",
    "open": "clasp open",
    "logs": "clasp logs"
  }
}
```

## Project Progress

Current status as of May 18, 2025:
- ✅ Project structure defined
- ✅ Critical functionality documented
- ✅ Migration plan created
- ✅ Skeleton implementations drafted
- ✅ Development environment configured
- 🔲 Core modules implementation (in progress)
- 🔲 Full functionality implementation
- 🔲 Testing and refinement
- 🔲 Final review and deployment

See `PROJECT_TASKS.md` for detailed task tracking.

## How to Resume Development with Claude Code

When continuing development with Claude Code:

1. **Start with the main project folder:**
   ```
   I'd like to continue working on the YSLv6Hub project. Please review the files in /Users/galagrove/yslv6hub/ to understand our project structure.
   ```

2. **Mention the backup documentation:**
   ```
   We've documented the critical GroupsTracker and SwimmerSkills functionality in detail in the backup folder at /Users/galagrove/Library/CloudStorage/GoogleDrive-ssullivan@penbayymca.org/My Drive/SRS YSLv6Hub/backup_for_claude/
   ```

3. **Reference where we left off:**
   ```
   We were working on implementing the modern development environment with CLASP, TypeScript, Jest, and ESLint, focusing on preserving the critical GroupsTracker/SwimmerSkills integration.
   ```

This will ensure that Claude has the full context of both the implementation code and the detailed documentation.

## Original Code Reference

The original code is available in two locations:
- `/Users/galagrove/YSLv6Hub-direct/`: Original direct export
- `/Users/galagrove/yslv6hub-gs/`: Organized legacy code for reference

These should be consulted when ensuring functionality matches the original implementation.

## Troubleshooting

### Common Issues

1. **CLASP Upload Errors**
   ```
   Error during upload: Insufficient permissions
   ```
   Solution: Re-authenticate with `clasp login` and ensure you have edit access to the Google Apps Script project.

2. **TypeScript Compilation Errors**
   ```
   Cannot find name 'SpreadsheetApp'
   ```
   Solution: Ensure proper Google Apps Script type definitions are included in `types/google-apps-script.d.ts`.

3. **Jest Test Failures with Google Apps Script Objects**
   ```
   ReferenceError: SpreadsheetApp is not defined
   ```
   Solution: Ensure proper mocks are set up in `test/mocks/` and imported in test files.

4. **Long-Running Scripts Timing Out**
   ```
   Execution time limit exceeded
   ```
   Solution: Implement chunking for large operations or use time-based triggers for batch processing.

## Contact

For support or questions, contact [ssullivan@penbayymca.org](mailto:ssullivan@penbayymca.org).

---

This project aims to improve the architecture and maintainability of the YSL application while ensuring that critical functionality, especially the GroupsTracker and SwimmerSkills integration, is preserved exactly as it functions in the original implementation.