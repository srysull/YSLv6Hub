# Testing YSLv6Hub Initialization Functionality

We've successfully deployed our YSLv6Hub code to Google Apps Script. Here's how to test the initialization functionality:

## Option 1: Test with the Linked Spreadsheet

1. Open the Google Spreadsheet we created:
   https://drive.google.com/open?id=1hAb9TgZaICZQUlJTG0buQFSa_NH5RD0WUVtnynUn71k

2. Wait for the spreadsheet to load - the onOpen trigger should run automatically

3. You should see a "YSLv6Hub" menu in the top menu bar
   - If you don't see it, refresh the page

4. Click on "YSLv6Hub" > "System" > "Test Initialization"

5. The initialization test should run and:
   - Create any missing sheets (YSLv6Hub, RegistrationInfo, SystemLog)
   - Set system properties
   - Show a dialog with the test results

6. Verify that:
   - All required sheets exist with proper formatting and headers
   - System properties are set correctly
   - The dialog shows all checks passing

## Option 2: Test with a Fresh Blank Spreadsheet

1. Go to https://sheets.google.com/ and create a new blank spreadsheet

2. In the spreadsheet menu, go to "Extensions" > "Apps Script"

3. In the Apps Script editor, click on "Project Settings" (gear icon)

4. Under the "Script ID" section, click "Change script ID"

5. Enter this script ID: `17mxN2QUfg6sWx7X88TYeJ_ceaxjp8g07b6MivqFzqnv0-u8Y60tEM9FV`

6. Confirm the change and reload the page

7. Go back to the spreadsheet and refresh the page

8. Follow steps 3-6 from Option 1 to test the initialization

## What to Look For

After running the initialization test, verify:

1. Three sheets are created:
   - YSLv6Hub - Dashboard with welcome text and instructions
   - RegistrationInfo - With column headers for student information
   - SystemLog - With column headers for logging

2. Each sheet has proper formatting:
   - Headers are bold and have a background color
   - Column widths are appropriately set
   - The YSLv6Hub dashboard has a title and instructions

3. The test results dialog shows:
   - ✅ All required sheets exist
   - ✅ System version: 6.0.0
   - ✅ Feature flags configured

## How to View the Code

You can view the code in the Google Apps Script editor:
https://script.google.com/d/17mxN2QUfg6sWx7X88TYeJ_ceaxjp8g07b6MivqFzqnv0-u8Y60tEM9FV/edit