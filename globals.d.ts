// Global declarations for Google Apps Script runtime

declare var GlobalFunctions: {
  onOpen: () => void;
  onEdit: (e: any) => void;
  extractIdFromUrl: (url: string) => string | null;
  safeGetFolderById: (id: string) => GoogleAppsScript.Drive.Folder | null;
  safeGetFileById: (id: string) => GoogleAppsScript.Drive.File | null;
  safeGetSpreadsheetById: (id: string) => GoogleAppsScript.Spreadsheet.Spreadsheet | null;
  formatStudentName: (firstName: string, lastName: string) => string;
  findColumnIndex: (headers: any[], possibleNames: string | string[]) => number;
  safeGetProperty: (key: string) => string | null;
  safeSetProperty: (key: string, value: string) => void;
  handleClassesSheetEdit: (e: any) => void;
  handleAnnouncementsSheetEdit: (e: any) => void;
  handleGroupLessonTrackerDropdownChange: (e: any) => void;
  populateStudentNames: (sheet: GoogleAppsScript.Spreadsheet.Sheet, classInfo: any) => void;
  populateClassDates: (sheet: GoogleAppsScript.Spreadsheet.Sheet, classInfo: any) => void;
  populateClassSkills: (sheet: GoogleAppsScript.Spreadsheet.Sheet, classInfo: any) => void;
  populateStudentSkills: (sheet: GoogleAppsScript.Spreadsheet.Sheet, classInfo: any) => void;
  syncStudentDataWithSwimmerSkills: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => void;
  collectStudentsFromGroupLessonTracker: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => any[];
  collectSkillsFromGroupLessonTracker: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => any;
  pullDataFromSwimmerSkills: (sheet: GoogleAppsScript.Spreadsheet.Sheet, students: any[], skills: any, swimmerSkillsData: any[][]) => void;
  getProgramAbbreviation: (program: string) => string;
  expandSheetForStudents: (sheet: GoogleAppsScript.Spreadsheet.Sheet, studentCount: number) => void;
  truncateSheetToOriginalSize: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => void;
  truncateSheetToSize: (sheet: GoogleAppsScript.Spreadsheet.Sheet, maxRow: number) => void;
  clearExistingTrackerData: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => void;
  applyAlternatingBackground: (sheet: GoogleAppsScript.Spreadsheet.Sheet, studentNames: any[]) => void;
  columnToLetter: (column: number) => string;
  testSyncFunctionality: () => void;
};

declare var ErrorHandling: {
  initializeErrorHandling: () => void;
  logMessage: (message: string, level: string, source: string) => void;
  handleError: (error: Error, functionName: string, userMessage: string, showAlert?: boolean) => void;
};

declare var AdministrativeModule: {
  getSystemConfiguration: () => any;
  generateInstructorSheets: () => boolean;
};

declare var VersionControl: {
  initializeVersionControl: () => void;
};

declare var CommunicationModule: {
  createCommunicationLog: () => void;
};

// Menu and UI functions
declare function createFixedMenu(): void;
declare function createFullMenu(): void;
declare function fixMenu(): void;
declare function installOnOpenTrigger(): void;
declare function reloadAllMenus(): void;
declare function fixSwimmerRecordsAccess(): void;
declare function directSyncStudentData(): void;

// Wrapper functions
declare function wrapInitializeBlankSpreadsheet(): void;
declare function showFixedConfigurationDialog(): void;
declare function wrapUpgradeExistingSpreadsheet(): void;