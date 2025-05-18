# Testing YSLv6Hub with a Blank Spreadsheet

## Testing Instructions

To test if the initialization functionality works correctly with a blank spreadsheet, follow these steps:

1. First, copy the script ID from our project:
   ```
   1jvHLWHyckIleHMWuNiCYLOrfQqBiaBogdP-4rCh1QEa5023Hoal-j1r_
   ```

2. Create a new blank Google Spreadsheet:
   - Go to https://sheets.google.com/ and create a new spreadsheet
   - Name it "YSLv6Hub Test"

3. Connect our script to the spreadsheet:
   - In the Google Sheets menu, click "Extensions" > "Apps Script"
   - In the Apps Script editor, go to "Project Settings" (gear icon)
   - Under the "Script ID" section, click "Change script ID"
   - Enter the script ID from step 1
   - Confirm the change

4. Test the initialization functionality:
   - Go back to the spreadsheet
   - Refresh the page to trigger the onOpen function
   - Check if the "YSLv6Hub" menu appears in the top menu bar
   - Click on "YSLv6Hub" > "System" > "Test Initialization"
   - Verify that the initialization test runs and shows the results dialog
   - Check that the required sheets (YSLv6Hub, RegistrationInfo, and SystemLog) are created automatically
   - Verify that each sheet has the proper formatting and headers

## Expected Results

After completing the test, you should observe:

1. The "YSLv6Hub" menu in the top menu bar with all the submenus
2. Three sheets should be created automatically:
   - YSLv6Hub (dashboard)
   - RegistrationInfo (with headers for student information)
   - SystemLog (with headers for logging)
3. The initialization test dialog should show:
   - ✅ All required sheets exist
   - ✅ System version: 6.0.0
   - ✅ Feature flags configured

## Troubleshooting

If the initialization doesn't work as expected:

1. Check the JavaScript console for any errors:
   - Open the browser developer tools (F12 or Ctrl+Shift+I)
   - Look for any error messages in the Console tab

2. Verify the script is correctly connected to the spreadsheet:
   - In Apps Script, check the Project Settings to confirm the script ID

3. Try manually running the initialization functions:
   - In Apps Script, add the following code to test:
   ```javascript
   function manualTest() {
     initializeSystem();
     createMenu();
   }
   ```
   - Run this function and check if it resolves the issues

4. If all else fails, try the repair function:
   - In the menu, go to "YSLv6Hub" > "System" > "Repair System"