# YSLv6Hub Enhanced

YSLv6Hub is a comprehensive Google Workspace system for managing the Youth Swim Lessons (YSL) program at PenBay YMCA. This enhanced version adds several key improvements to make the system more robust, user-friendly, and feature-rich.

## Key Features

### Core Functionality
- Class management and instructor resources
- Student roster management and assessment tracking
- Email communications with parents and participants
- Reporting on student progress
- System administration and configuration

### New Enhancements
- **Fixed Configuration Dialog**: Properly handles blank spreadsheets and provides a more robust configuration interface
- **Blank Spreadsheet Initializer**: Creates all required sheets and structure from a completely blank spreadsheet
- **Email Templates System**: Reusable email templates with placeholders for personalization
- **Input Validation**: Enhanced data validation throughout the system for improved data integrity

## Getting Started

### For New Spreadsheets
1. Open a new or blank Google Sheet
2. Go to Extensions > Apps Script
3. Copy all `.gs` files from this repository into the Apps Script editor
4. Save and close the editor
5. Refresh your spreadsheet
6. Use the "YSL Hub Enhanced" menu > "Initialize Blank Spreadsheet"
7. Follow the initialization wizard

### For Existing YSL Hub Spreadsheets
1. Go to Extensions > Apps Script
2. Copy all `.gs` files from this repository into the Apps Script editor
3. Save and close the editor
4. Refresh your spreadsheet
5. Use the "YSL Hub Enhanced" menu > "Run Full Upgrade"

## Module Structure

- **AdministrativeModule.gs**: System initialization and configuration
- **CommunicationModule.gs**: Email and notifications
- **DataIntegrationModule.gs**: Data processing and management
- **ErrorHandling.gs**: Centralized error handling and logging
- **Globals.gs**: Global functions and utilities
- **InstructorResourceModule.gs**: Instructor-specific tools
- **MenuWrappers.gs**: Menu creation and event handlers
- **ReportingModule.gs**: Assessment reports generation
- **VersionControl.gs**: Version management and updates

### New Modules
- **BlankSheetInitializer.gs**: Tools for creating system structure in a blank spreadsheet
- **EmailTemplates.gs**: Template management for communication
- **FixedConfigDialog.gs**: Fixed version of configuration dialog
- **InputValidation.gs**: Data validation functions and utilities
- **YSLv6HubUpgrade.gs**: Integration of all enhancements

## Using Email Templates

The enhanced email system allows you to create and manage templates for common communications:

1. Access through "YSL Hub Enhanced" menu > "Email Templates" > "Manage Email Templates"
2. Create new templates with placeholders (e.g., `{student_name}`, `{class_day}`)
3. Use templates when sending emails to participants

### Default Templates
- Welcome Email
- Lesson Reminder
- Assessment Completed
- Class Cancellation
- Session Ending Summary

## Data Validation

Input validation ensures data integrity across the system:

- Validates email addresses, phone numbers, dates, and times
- Provides dropdown lists for common fields (days of the week, status values)
- Shows user-friendly error messages when validation fails

## Next Steps and Future Enhancements

Planned improvements for future versions:
- Reporting dashboard for instructors
- Mobile-friendly UI improvements
- Enhanced error handling with more user-friendly messages
- Performance optimizations for large rosters

## Support

For support or to report issues, please contact:
- ssullivan@penbayymca.org

## Version

YSLv6Hub Enhanced v2.1.0
Last Updated: May 5, 2025