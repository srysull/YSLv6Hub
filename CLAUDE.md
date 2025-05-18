# YSL v6 Hub Project Notes

## Project Overview
- **Name**: YSL Swim Lessons Hub
- **Purpose**: Google Sheets-based application for managing swim lesson tracking
- **Tech Stack**: Google Apps Script, TypeScript, CLASP

## Key Components
- **Group Lesson Tracker**: Generated sheet for instructors to track student skills
- **SwimmerSkills**: Sheet containing all student skills data
- **Menu System**: Centralized in `00_MenuSystem.ts` to avoid conflicts

## Recent Fixes

### 1. Menu System Improvements
- Centralized menu system in `00_MenuSystem.ts` 
- Fixed duplicate `onOpen` functions
- Simplified to a single menu with clear organization
- Added test functions to Tools & Diagnostics submenu

### 2. Git/Google Drive Integration
- Created `sync-ysl-folders.sh` script to sync repository with Google Drive folder
- Implemented git hooks to automate syncing
- Set git author information for Sean Sullivan

### 3. Group Lesson Tracker Enhancements
- Added confirmation prompt before replacing existing sheets
- Fixed student data sync to prevent overwriting source data
- Improved UI with clearer messages

### 4. Sync Student Data Fix (Latest)
- Fixed `syncStudentDataWithSwimmerSkills` function in `01_Globals.ts`
- Now correctly writes to "Repeat" column (one column to the right) in SwimmerSkills
- Added better error handling and improved column detection
- Added test function and menu item for easy testing

## Important Files
- `00_MenuSystem.ts`: Main menu creation and management
- `00_TriggerFunctions.ts`: Event handlers for main triggers
- `01_Globals.ts`: Core functions including sync functionality
- `15_DynamicInstructorSheet.ts`: Group Lesson Tracker generation

## Commands to Run During Development
- `clasp push`: Push changes to Google Apps Script project
- `./sync-ysl-folders.sh`: Manually sync to Google Drive if needed

## Testing
Use the "Test Sync Functionality" option in the Tools & Diagnostics menu to verify the sync functionality.