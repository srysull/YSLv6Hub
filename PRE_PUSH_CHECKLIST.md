# Pre-Push Testing Checklist

## 🚨 MANDATORY: Complete ALL items before pushing to Google Apps Script

### 1. Menu Display Verification
- [ ] Run `clasp push` to deploy
- [ ] Open the Google Spreadsheet
- [ ] **RELOAD the spreadsheet** (Ctrl+R or Cmd+R)
- [ ] **VERIFY: YSL v6 Hub menu appears in the menu bar**
- [ ] If menu doesn't appear:
  - [ ] Open Apps Script editor
  - [ ] Run `checkOnOpenTriggers()` to diagnose
  - [ ] Run `installOnOpenTrigger()` to fix
  - [ ] Reload spreadsheet again to verify fix

### 2. Core Functionality Tests
- [ ] Click "Generate Group Lesson Tracker" - verify it opens dialog
- [ ] Click "Sync Student Data" - verify it runs without error
- [ ] Click "System Configuration" - verify dialog opens
- [ ] Test at least one submenu item from each category

### 3. Error Handling Tests
- [ ] Check browser console for JavaScript errors (F12)
- [ ] Check Apps Script logs for errors
- [ ] Verify error messages are user-friendly

### 4. TypeScript Compilation
- [ ] Run `npm run typecheck` - must pass with 0 errors
- [ ] Run `npm test` - all unit tests must pass
- [ ] Check for TypeScript strict mode violations

### 5. Deployment Verification
- [ ] Verify `.claspignore` excludes test files and duplicates
- [ ] Check `clasp status` shows correct files to push
- [ ] Ensure no duplicate file names will be pushed

### 6. Git Commit Standards
- [ ] Stage all relevant files
- [ ] Write descriptive commit message
- [ ] Include what was fixed/added
- [ ] Push to GitHub after successful GAS deployment

## 🔴 CRITICAL: Menu Must Display Test

**THE MOST IMPORTANT TEST**: After deployment, close the spreadsheet completely and reopen it. The menu MUST appear without any manual intervention.

If the menu doesn't appear automatically:
1. DO NOT push to Git
2. DO NOT mark the task as complete
3. Debug and fix the issue first
4. Re-run this entire checklist

## Automated Test Command (Future Implementation)

```bash
# Run this before every push:
npm run pre-push-test
```

This will:
1. Compile TypeScript
2. Run unit tests
3. Deploy to test environment
4. Verify menu displays
5. Run basic functionality tests
6. Report success/failure

## Emergency Fix Procedures

If menu fails after deployment:
1. Check if `onOpen` function exists in Apps Script editor
2. Verify function is at global scope (not inside another function)
3. Check for JavaScript syntax errors
4. Try creating a simple test: `function onOpen() { SpreadsheetApp.getUi().alert('Test'); }`
5. If test works, gradually add menu items back

## Known Issues to Check

- [ ] TypeScript compilation not preserving function names
- [ ] Clasp deployment order causing issues
- [ ] Multiple onOpen functions conflicting
- [ ] Trigger not being installed automatically
- [ ] Google Apps Script caching issues

---

**Remember**: A non-functional menu makes the entire system unusable. This checklist is not optional!