# YSLv6Hub Testing Strategy

This document outlines the comprehensive testing strategy for the refactored YSLv6Hub system. It combines both manual testing for Google Sheets functionality and automated testing using Jest for the TypeScript codebase.

## Testing Environment Setup

### Local Development Testing

1. **Unit Testing with Jest**
   - All TypeScript code is tested using Jest with Google Apps Script service mocks
   - Run tests locally with: `npm test`
   - Use watch mode during development: `npm test -- --watch`

2. **Static Analysis**
   - ESLint is configured to enforce code quality standards
   - Run linting with: `npm run lint`
   - TypeScript compilation checks: `npm run typecheck`

3. **Pre-commit Hooks**
   - All tests and linting run automatically before each commit
   - Failed tests or linting issues prevent commits

### Google Sheets Testing

1. **Test Deployment**
   - Deploy code to a test Google Sheet using CLASP
   - Command: `npm run deploy:test`
   - This deploys to a designated test sheet without affecting production

2. **Testing Environment Setup**
   - The test sheet is pre-populated with sample data
   - Test accounts are set up with various permission levels

## Testing Levels

### 1. Unit Testing

Unit tests verify individual functions in isolation:

- **Core Module Tests**: Verify utility functions, error handling, and logging
- **Data Access Tests**: Validate sheet reading/writing operations
- **Template Tests**: Confirm UI generation works correctly
- **GroupsTracker Tests**: Verify sheet creation and formatting
- **SkillsSync Tests**: Validate bidirectional synchronization logic

Each module has its own test file with complete coverage of all exported functions.

### 2. Integration Testing

Integration tests verify that modules work together:

- **GroupsTracker + SkillsSync Integration**: Verify bidirectional sync works end-to-end
- **UI + Core Integration**: Test that UI operations correctly trigger core functionality
- **Menu + Module Integration**: Confirm menu options correctly invoke the right modules

### 3. End-to-End Testing

End-to-end tests verify complete workflows:

- **Full GroupsTracker Creation Workflow**: From setup to student data import
- **Complete Sync Cycle**: Testing both directions of synchronization
- **Error Recovery**: Verifying that the system recovers from various failure scenarios

## Test Cases for Critical Functionality

### 1. GroupsTracker Generation

#### 1.1 Sheet Creation and Formatting
- [ ] Generate GroupsTracker sheet for a new class
- [ ] Verify header formatting is correctly applied
- [ ] Confirm all columns have proper headers
- [ ] Check that sheet protection is correctly applied

#### 1.2 Student Data Population
- [ ] Import student data from RegistrationInfo sheet
- [ ] Verify all student records are imported correctly
- [ ] Check that student names, ages, and other details match the source data
- [ ] Test with varying class sizes (small, medium, large)

#### 1.3 Skills Column Generation
- [ ] Verify skill columns are created for the appropriate level
- [ ] Check that skill descriptions are correct
- [ ] Confirm that skill order matches the curriculum

### 2. Bidirectional Synchronization

#### 2.1 GroupsTracker to SwimmerSkills
- [ ] Update skills data in GroupsTracker
- [ ] Run sync operation
- [ ] Verify changes are reflected in SwimmerSkills sheet
- [ ] Test with various data types (text, numbers, dates)

#### 2.2 SwimmerSkills to GroupsTracker
- [ ] Update student data in SwimmerSkills
- [ ] Run sync operation
- [ ] Verify changes are reflected in GroupsTracker sheet
- [ ] Test with various skill ratings and comments

#### 2.3 Conflict Resolution
- [ ] Create deliberate conflicts by changing the same data in both sheets
- [ ] Run sync operation
- [ ] Verify conflict resolution policy is correctly applied
- [ ] Check that newer changes take precedence

### 3. Error Handling and Recovery

#### 3.1 Input Validation
- [ ] Test with invalid skill ratings
- [ ] Verify validation errors are appropriately displayed
- [ ] Confirm that invalid data is not synchronized

#### 3.2 Service Interruptions
- [ ] Simulate API limits and service failures
- [ ] Verify the system handles these gracefully
- [ ] Check that user is informed with appropriate error messages
- [ ] Confirm that the system offers recovery options

#### 3.3 Edge Cases
- [ ] Test with very large classes
- [ ] Verify handling of special characters in names and comments
- [ ] Test with empty sheets and missing data

## Test Reporting

For each test execution:

1. **Automated Test Reports**
   - Jest outputs detailed test results including coverage reports
   - Failed tests include stack traces and error messages

2. **Manual Test Documentation**
   - For each manual test case, record:
     - Pass/Fail status
     - Screenshots of relevant issues
     - Steps to reproduce any failures
     - Environment details (browser, OS, etc.)

## Continuous Integration

Whenever code is pushed to the repository:

1. All automated tests are run
2. Code quality checks are performed
3. Build verification tests ensure the code compiles correctly
4. Test coverage reports are generated

## Known Issues and Limitations

Any known issues or testing limitations should be documented here:

1. *Google Apps Script quotas may limit the number of operations in a single execution*
2. *Testing multiple concurrent users requires manual coordination*

## Test Data Management

Test data is managed through:

1. **Seed Scripts**: Automatically populate test sheets with known data
2. **Test Fixtures**: JSON files containing sample data for various scenarios
3. **Reset Functions**: Quickly restore sheets to a known state between tests

---

*Last Updated: May 18, 2025*