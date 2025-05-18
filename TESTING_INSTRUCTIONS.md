# YSLv6Hub Testing Instructions

To test the initialization functionality with a blank Google workbook, follow these step-by-step instructions:

## Creating a New Google Apps Script Project

1. Go to https://script.google.com/
2. Click "New Project"
3. Rename the project to "YSLv6Hub"

## Setting Up Google Apps Script Files

1. Copy and paste the following code files into the Apps Script editor:

### 00_System.js
- Copy the content from `/Users/galagrove/yslv6hub/build/00_System.js`
- Create a new file in the Apps Script editor named "00_System"
- Paste the code and save

### 01_Core.js
- Copy the content from `/Users/galagrove/yslv6hub/build/01_Core.js`
- Create a new file in the Apps Script editor named "01_Core"
- Paste the code and save

### 10_SystemLog.js
- Copy the content from `/Users/galagrove/yslv6hub/build/10_SystemLog.js`
- Create a new file in the Apps Script editor named "10_SystemLog"
- Paste the code and save

### interfaces/index.js
- Copy the content from `/Users/galagrove/yslv6hub/build/interfaces/index.js`
- Create a new file in the Apps Script editor named "interfaces_index"
- Paste the code and save

### utils/constants.js
- Copy the content from `/Users/galagrove/yslv6hub/build/utils/constants.js`
- Create a new file in the Apps Script editor named "utils_constants"
- Paste the code and save

### appsscript.json
- Copy the content from `/Users/galagrove/yslv6hub/appsscript.json`
- In the Apps Script editor, click on "Project Settings"
- Under the "Script Properties" section, add the appropriate values from the JSON file

## Creating a Test Spreadsheet

1. Go to https://sheets.google.com/
2. Create a new blank spreadsheet
3. Rename it to "YSLv6Hub Test"

## Connecting the Script to the Spreadsheet

1. In the Google Sheets menu, click "Extensions" > "Apps Script"
2. Delete any default code in the editor
3. Copy and paste the same files as you did for the standalone project
4. Save all files

## Testing Initialization

1. Go back to your spreadsheet
2. Refresh the page to trigger the onOpen function
3. A new menu item "YSLv6Hub" should appear in the top menu
4. Click on "YSLv6Hub" to expand the menu
5. Under the "System" submenu, click on "About" to verify the menu works
6. Create a custom function to test initialization:
   - In the Apps Script editor, create a function that calls `testInitialization()`
   - Run this function directly from the Apps Script editor
   - This will test if all required sheets are created and system properties are set

## Expected Results

After refreshing the spreadsheet, you should see:

1. A "YSLv6Hub" menu item in the top menu bar
2. If you click on the menu, all submenu items should be visible
3. When you run the `testInitialization()` function, you should see:
   - A popup showing test results
   - Three sheets should be created automatically: "YSLv6Hub", "RegistrationInfo", "SystemLog"
   - Each sheet should have the proper formatting and headers
   - System properties should be initialized

## Testing Error Recovery

To test the error recovery mechanism:

1. Deliberately introduce an error in the `initializeSystem()` function
2. Refresh the spreadsheet
3. You should see the emergency menu appear instead of the regular menu
4. Click on "Repair System" in the emergency menu
5. The system should recover and create the proper menu and sheets

## Testing Logs

To verify that logging works correctly:

1. Look at the SystemLog sheet
2. After initialization, there should be log entries for the initialization process
3. Use the Apps Script logs to see console messages (View > Logs)

## Debugging Tips

- If menus don't appear, check the Apps Script execution logs for errors
- Make sure all referenced functions are properly exposed through the global object
- Check that all files are correctly imported and in the right order
- Verify that the appsscript.json file has the correct configuration