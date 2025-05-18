# CLAUDE.md - Instructions for Claude Code

This file contains instructions and reminders for Claude Code when working on the YSLv6Hub project.

## Project Overview

YSLv6Hub is a TypeScript refactoring of a YMCA Swim Lessons Google Sheets application, providing improved organization, type safety, and modern development practices while preserving critical functionality. It streamlines the management of swim lessons with features for registration data handling, class tracking, skills assessment, attendance logging, communications, and reporting.

## Critical Functionality

The most crucial functionality that must be preserved is the **GroupsTracker and SwimmerSkills integration**, specifically:

1. GroupsTracker sheet generation with formatted templates
2. Student data population from roster information
3. Skills population based on class level/stage
4. Bidirectional synchronization between GroupsTracker and SwimmerSkills sheets

## Development Environment Commands

When working on this project, use these commands:

```bash
# Install dependencies
npm install

# Lint code
npm run lint

# Type check
npm run typecheck

# Run tests
npm test

# Build the project
npm run build

# Deploy to Google Apps Script
npm run deploy
```

## Project Structure

- **src/**: Source TypeScript files
  - **00_System.ts**: Main entry point, initialization, menu creation
  - **01_Core.ts**: Core utilities, error handling, caching, events
  - **02_DataAccess.ts**: Sheet access, data retrieval, manipulation
  - **03_Templates.ts**: UI templates and components
  - **04_GroupsTracker.ts**: GroupsTracker sheet creation and management
  - **05_SkillsSync.ts**: Bidirectional sync with SwimmerSkills
  - Additional modules as needed
- **tests/**: Jest test files
- **build/**: Compiled JavaScript output (for CLASP)
- **types/**: Custom type definitions
- **tools/**: Development utility scripts

## Google Apps Script Limitations

Be aware of these constraints when implementing:
- Max execution time: 6 minutes per execution
- Max spreadsheet operations: ~30,000 cells per minute
- No ES modules: Use namespace approach instead
- Limited modern JS features: Target ES2019 compatibility

## Bidirectional Sync Behavior

The critical bidirectional sync must function exactly as before:
- GroupsTracker "End" columns → SwimmerSkills "Repeat" columns (one column to right)
- SwimmerSkills columns → GroupsTracker "Beginning" columns
- Original SwimmerSkills data must be preserved
- Color coding: X = green (completed), / = yellow (taught)

## Error Handling Framework

All functions must use the standardized error handling:
```typescript
try {
  // Function logic
} catch (error) {
  Core.handleError(error, 'ModuleName.functionName', 'User-friendly error message');
  return null; // Or appropriate error indicator
}
```

## Additional Documentation

For detailed project information, refer to:
- README.md - Main project documentation
- PROJECT_TASKS.md - Detailed task tracking
- PROJECT_STATUS.md - Current status and next steps
- TESTING.md - Testing strategy and test cases

## Backup Documentation Location

Important functionality documentation is also available at:
`/Users/galagrove/Library/CloudStorage/GoogleDrive-ssullivan@penbayymca.org/My Drive/SRS YSLv6Hub/backup_for_claude/`