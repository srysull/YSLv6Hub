# YSLv6Hub Enhanced - Testing Plan

This document outlines the testing process for the enhanced YSLv6Hub system. Use this plan to verify that all new features work correctly.

## Test Environment Setup

1. Create a new Google Sheet
2. Go to Extensions > Apps Script
3. Copy all numbered `.gs` files (01_Globals.gs through 17_HistoryModule.gs) and FixedConfigDialog.gs into the Apps Script editor
4. Save and close the editor
5. Refresh your spreadsheet

## Test Plan

### 1. System Initialization & Configuration

#### 1.1 Blank Spreadsheet Initialization
- [ ] Access "YSL Hub Setup" menu
- [ ] Run "Initialize Blank Spreadsheet"
- [ ] Verify all required sheets are created (Assumptions, Classes, Roster, etc.)
- [ ] Confirm basic formatting is applied to each sheet

#### 1.2 Fixed Configuration Dialog
- [ ] Access "YSL Hub Enhanced" menu
- [ ] Run "System Configuration (Fixed)"
- [ ] Verify dialog displays without errors
- [ ] Enter test configuration values
- [ ] Apply changes and verify they are saved

### 2. Email Templates System

#### 2.1 Template Management
- [ ] Access "Email Templates" > "Manage Email Templates"
- [ ] Verify default templates are available
- [ ] Create a new template with test placeholders
- [ ] Edit an existing template
- [ ] Verify changes are saved

#### 2.2 Templated Emails
- [ ] Create a test class with at least one student
- [ ] Select the test class
- [ ] Access "Email Templates" > "Send Templated Email"
- [ ] Select a template and send a test email
- [ ] Verify placeholders are correctly replaced
- [ ] Verify email is received at the destination

### 3. Input Validation

#### 3.1 Sheet Validation Rules
- [ ] Run "Apply Sheet Validation"
- [ ] Check Classes sheet for dropdown lists (days of week, Select/Exclude)
- [ ] Check Announcements sheet for status dropdown (Draft, Ready, Sent, Failed)
- [ ] Check Assessments sheet for rating dropdown

#### 3.2 Data Validation
- [ ] Enter invalid email format in Roster sheet
- [ ] Enter invalid date format
- [ ] Enter out-of-range numeric values
- [ ] Verify validation errors are shown

### 4. Integration & Upgrade

#### 4.1 Full Upgrade Process
- [ ] In a copy of an existing YSL Hub spreadsheet
- [ ] Run "Run Full Upgrade" function
- [ ] Verify upgrade summary is shown
- [ ] Check that Upgrade Guide sheet is created
- [ ] Confirm enhanced menu is available

#### 4.2 Cross-Module Integration
- [ ] Verify Email Templates can be used with existing Class Management
- [ ] Verify Input Validation works with existing data entry
- [ ] Test that configuration changes are properly validated

## Regression Testing

### 5.1 Core Functionality
- [ ] Create a class and add students
- [ ] Generate instructor sheets
- [ ] Send a basic email to class participants
- [ ] Generate a report
- [ ] Verify all core functions continue to work as expected

### 5.2 Error Handling
- [ ] Trigger intentional errors (invalid URLs, missing data)
- [ ] Verify that the error handling system captures and logs errors
- [ ] Check that user-friendly error messages are displayed

## Test Reporting

For each test item, record:
- Pass/Fail status
- Any unexpected behavior
- Error messages encountered
- Browser and device used

## Completion Criteria

Testing is considered complete when:
1. All test cases have been executed
2. All critical issues have been resolved
3. The system works correctly on both new and existing spreadsheets

## Known Issues

Document any known issues or limitations discovered during testing:

1. *Issue description, conditions, workaround (if any)*
2. *...*