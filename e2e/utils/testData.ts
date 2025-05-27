/**
 * Test Data Management Utilities
 */

import { SheetsHelper } from './sheets';

export interface TestClass {
  name: string;
  day: string;
  time: string;
  instructor?: string;
  location?: string;
}

export interface TestStudent {
  firstName: string;
  lastName: string;
  email: string;
  className: string;
  level?: string;
}

export interface TestAssumptions {
  [key: string]: string | number;
}

export class TestDataHelper {
  constructor(private sheets: SheetsHelper) {}

  /**
   * Set up default test assumptions
   */
  async setupAssumptions(assumptions?: TestAssumptions): Promise<void> {
    const defaults: TestAssumptions = {
      'Session Duration': 30,
      'Instructor Name': 'Test Instructor',
      'Pool Name': 'Test Pool',
      'Organization': 'Test YMCA',
      'Email From': 'test@example.com',
      ...assumptions
    };

    await this.sheets.switchToSheet('Assumptions');
    
    let row = 2;
    for (const [key, value] of Object.entries(defaults)) {
      await this.sheets.setCellValue(`A${row}`, key);
      await this.sheets.setCellValue(`B${row}`, String(value));
      row++;
    }
  }

  /**
   * Add test classes
   */
  async addTestClasses(classes: TestClass[]): Promise<void> {
    await this.sheets.switchToSheet('Classes');
    
    let row = 2;
    for (const classData of classes) {
      await this.sheets.setCellValue(`A${row}`, classData.name);
      await this.sheets.setCellValue(`B${row}`, classData.day);
      await this.sheets.setCellValue(`C${row}`, classData.time);
      if (classData.instructor) {
        await this.sheets.setCellValue(`D${row}`, classData.instructor);
      }
      if (classData.location) {
        await this.sheets.setCellValue(`E${row}`, classData.location);
      }
      row++;
    }
  }

  /**
   * Add test students
   */
  async addTestStudents(students: TestStudent[]): Promise<void> {
    await this.sheets.switchToSheet('Roster');
    
    let row = 2;
    for (const student of students) {
      await this.sheets.setCellValue(`A${row}`, student.firstName);
      await this.sheets.setCellValue(`B${row}`, student.lastName);
      await this.sheets.setCellValue(`C${row}`, student.email);
      await this.sheets.setCellValue(`D${row}`, student.className);
      if (student.level) {
        await this.sheets.setCellValue(`E${row}`, student.level);
      }
      row++;
    }
  }

  /**
   * Generate sample test data
   */
  static generateSampleData() {
    const classes: TestClass[] = [
      { name: 'Beginner Swim 101', day: 'Monday', time: '10:00 AM', instructor: 'John Smith' },
      { name: 'Intermediate Swim 201', day: 'Wednesday', time: '2:00 PM', instructor: 'Jane Doe' },
      { name: 'Advanced Swim 301', day: 'Friday', time: '4:00 PM', instructor: 'Bob Johnson' }
    ];

    const students: TestStudent[] = [
      { firstName: 'Alice', lastName: 'Anderson', email: 'alice@example.com', className: 'Beginner Swim 101', level: 'Beginner' },
      { firstName: 'Bob', lastName: 'Brown', email: 'bob@example.com', className: 'Beginner Swim 101', level: 'Beginner' },
      { firstName: 'Charlie', lastName: 'Chen', email: 'charlie@example.com', className: 'Intermediate Swim 201', level: 'Intermediate' },
      { firstName: 'Diana', lastName: 'Davis', email: 'diana@example.com', className: 'Advanced Swim 301', level: 'Advanced' },
      { firstName: 'Edward', lastName: 'Evans', email: 'edward@example.com', className: 'Intermediate Swim 201', level: 'Intermediate' }
    ];

    const assumptions: TestAssumptions = {
      'Session Duration': 45,
      'Instructor Name': 'Head Instructor',
      'Pool Name': 'Main Pool',
      'Organization': 'Community YMCA',
      'Email From': 'swim@ymca.org',
      'Term Start Date': '2024-01-15',
      'Term End Date': '2024-03-15',
      'Number of Sessions': 8
    };

    return { classes, students, assumptions };
  }

  /**
   * Clear all data from sheets
   */
  async clearAllData(): Promise<void> {
    const sheetsToClean = ['Assumptions', 'Classes', 'Roster'];
    
    for (const sheetName of sheetsToClean) {
      try {
        await this.sheets.switchToSheet(sheetName);
        // Select all data rows (keeping headers)
        await this.sheets.selectRange('A2:Z1000');
        await this.sheets.pressKey('Delete');
        await new Promise(resolve => setTimeout(resolve, 500));
      } catch (error) {
        console.log(`Could not clear ${sheetName}: ${error.message}`);
      }
    }
  }

  /**
   * Verify data was saved correctly
   */
  async verifyDataIntegrity(): Promise<{success: boolean, errors: string[]}> {
    const errors: string[] = [];
    
    // Check Classes sheet
    await this.sheets.switchToSheet('Classes');
    const firstClassName = await this.sheets.getCellValue('A2');
    if (!firstClassName) {
      errors.push('No classes found in Classes sheet');
    }
    
    // Check Roster sheet
    await this.sheets.switchToSheet('Roster');
    const firstStudentName = await this.sheets.getCellValue('A2');
    if (!firstStudentName) {
      errors.push('No students found in Roster sheet');
    }
    
    // Check Assumptions sheet
    await this.sheets.switchToSheet('Assumptions');
    const firstAssumption = await this.sheets.getCellValue('A2');
    if (!firstAssumption) {
      errors.push('No assumptions found in Assumptions sheet');
    }
    
    return {
      success: errors.length === 0,
      errors
    };
  }
}