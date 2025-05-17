# YSLv6Hub TypeScript Project

This is a TypeScript implementation of the YSL Hub system for managing YMCA swim lessons, using Google Apps Script.

## Project Structure

This project uses clasp's direct TypeScript compilation capability to deploy TypeScript files directly to Google Apps Script.

Key files:
- `00_TriggerFunctions.ts` - Entry points for menu and edit triggers
- `01_Globals.ts` - Common functions and utilities
- `02_ErrorHandling.ts` - Error handling and logging
- `03_VersionControl.ts` - Version control and diagnostics
- `04_AdministrativeModule.ts` - System administration
- `05_MenuWrappers.ts` - Menu function wrappers
- `15_DynamicInstructorSheet_Skills.ts` - Skills functions

## Development Workflow

1. Edit the TypeScript files in this directory
2. Deploy to Google Apps Script with `clasp push`
3. Access your project in the Google Apps Script editor

## Accessing Your Google Apps Script Project

Visit your project directly at:
https://script.google.com/home/projects/17mxN2QUfg6sWx7X88TYeJ_ceaxjp8g07b6MivqFzqnv0-u8Y60tEM9FV/edit

## TypeScript Guidelines

When writing TypeScript for this project:

1. Do not use type annotations in function parameters or return types
2. Do not use interfaces or type declarations
3. Do not use imports or exports
4. Write standard Google Apps Script code with TypeScript features like:
   - Arrow functions
   - Template literals
   - Optional chaining
   - Destructuring
   - Default parameters
   
## Future Improvements

- Create a build process to compile TypeScript with full type checking locally
- Implement unit tests
- Add more advanced TypeScript features with automatic type removal for deployment