# YSLv6Hub Testing with a Blank Spreadsheet

Since we're having technical issues with CLASP, follow these manual steps to test the initialization functionality:

## Step 1: Open the Google Apps Script Project

1. Go to this URL:
   https://script.google.com/d/1SrgzOffSITbpM5_qMiVN1rRp_0vEg2fF0_zBJ0A6t4zbboaE6K_eySu7/edit

2. You should see an empty project or a project with just appsscript.json

## Step 2: Add the System Module

1. In the Apps Script editor, click on the "+" next to "Files" to create a new file
2. Name it "Code" (this is the default entry point)
3. Copy and paste the content from:
   `/Users/galagrove/yslv6hub/dist/00_System.js`

## Step 3: Test with a Blank Spreadsheet

1. From the Apps Script editor, click on "Deploy" > "New deployment"
2. Select type: "Web app"
3. Give it a description: "YSLv6Hub Test"
4. Set "Execute as" to "User accessing the web app"
5. Set "Who has access" to "Anyone"
6. Click "Deploy"
7. Copy the Web app URL

8. Create a new blank Google Spreadsheet:
   - Go to https://sheets.google.com/ and create a new spreadsheet
   - Name it "YSLv6Hub Test"

9. In the spreadsheet, go to "Extensions" > "Apps Script"
10. Delete any code in the default file
11. Click on "Project Settings" (gear icon)
12. Under "Script ID", click "Change script ID"
13. Enter: `1SrgzOffSITbpM5_qMiVN1rRp_0vEg2fF0_zBJ0A6t4zbboaE6K_eySu7`
14. Confirm the change
15. Reload the page

## Step 4: Verify Initialization

1. Refresh the spreadsheet
2. You should see a "YSLv6Hub" menu in the top menu bar
3. Click on "YSLv6Hub" > "System" > "Test Initialization"
4. Verify that the initialization test runs and creates all required sheets:
   - YSLv6Hub (dashboard)
   - RegistrationInfo (with headers for student information)
   - SystemLog (with headers for logging)
5. The test should show:
   - ✅ All required sheets exist
   - ✅ System version: 6.0.0
   - ✅ Feature flags configured

## Expected Results

After completing the test, verify:

1. The "YSLv6Hub" menu appears in the top menu bar
2. Three sheets are created automatically: YSLv6Hub, RegistrationInfo, and SystemLog
3. Each sheet has the correct formatting and headers
4. The initialization test dialog shows all checks passing

## Alternative Method if Script ID Change Doesn't Work

If you can't change the script ID:

1. Create a new blank Google Spreadsheet
2. Go to "Extensions" > "Apps Script"
3. Create a new file called "Code"
4. Copy and paste the code from `/Users/galagrove/yslv6hub/dist/00_System.js`
5. Save the project
6. Reload the spreadsheet
7. Test as described in Step 4