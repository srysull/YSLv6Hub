/**
 * Custom type definitions for Google Apps Script
 * Supplements the @types/google-apps-script package
 */

declare namespace GoogleAppsScript {
  /**
   * Custom interfaces and types for YSLv6Hub
   */
  namespace YSLv6Hub {
    /**
     * Student interface for tracking student information
     */
    interface Student {
      id: string;
      firstName: string;
      lastName: string;
      stage: string;
      age: number;
      guardian: string;
      email: string;
      phone: string;
      notes?: string;
    }

    /**
     * Class interface for tracking class information
     */
    interface Class {
      id: string;
      name: string;
      stage: string;
      instructorName: string;
      schedule: string;
      location: string;
      startDate: Date;
      endDate: Date;
      studentIds: string[];
    }

    /**
     * Skill interface for tracking student skills
     */
    interface Skill {
      id: string;
      name: string;
      description: string;
      stage: string;
      category: string;
      order: number;
    }

    /**
     * Student skill assessment
     */
    interface SkillAssessment {
      studentId: string;
      skillId: string;
      status: 'not-introduced' | 'introduced' | 'developing' | 'completed';
      assessedOn?: Date;
      notes?: string;
    }
  }
}