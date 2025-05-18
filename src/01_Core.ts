/**
 * YSLv6Hub Core Module
 * 
 * This module provides core utilities, error handling, event management, and caching.
 * It serves as the foundation for all other modules in the system.
 * 
 * @author Sean R. Sullivan
 * @version 1.0.0
 * @date 2025-05-18
 */

/**
 * Error severity levels
 */
export enum ErrorSeverity {
  INFO = 'info',
  WARNING = 'warning',
  ERROR = 'error',
  CRITICAL = 'critical'
}

/**
 * Error context information
 */
export interface ErrorContext {
  module: string;
  function: string;
  userMessage?: string;
  technicalDetails?: Record<string, any>;
}

/**
 * Feature flags for enabling/disabling features
 */
export enum FeatureFlag {
  SMART_IMPORT = 'SMART_IMPORT',
  FORM_PROCESSOR = 'FORM_PROCESSOR',
  ENHANCED_COMMUNICATIONS = 'ENHANCED_COMMUNICATIONS',
  ADVANCED_REPORTING = 'ADVANCED_REPORTING',
  SESSION_ANALYTICS = 'SESSION_ANALYTICS'
}

/**
 * User roles for access control
 */
export enum UserRole {
  ADMIN = 'ADMIN',
  COORDINATOR = 'COORDINATOR',
  INSTRUCTOR = 'INSTRUCTOR',
  VIEWER = 'VIEWER'
}

/**
 * User experience levels for progressive disclosure
 */
export enum UserExperienceLevel {
  BEGINNER = 'BEGINNER',
  INTERMEDIATE = 'INTERMEDIATE',
  ADVANCED = 'ADVANCED'
}

/**
 * Cache options interface
 */
export interface CacheOptions {
  expirySeconds?: number;
}

/**
 * ErrorHandling system for centralized error management
 */
export const ErrorHandling = {
  /**
   * Handles an error with proper logging and notification
   */
  handleError(error: Error, context: ErrorContext, severity: ErrorSeverity = ErrorSeverity.WARNING): void {
    // Log to console
    console.error(`[${severity}] ${context.module}.${context.function}: ${error.message}`);
    
    // Log to SystemLog sheet
    this.logToSystemLog(error, context, severity);
    
    // Notify user for higher severity errors
    if (severity === ErrorSeverity.ERROR || severity === ErrorSeverity.CRITICAL) {
      this.notifyUser(context.userMessage || error.message, severity);
    }
    
    // Report critical errors to admin
    if (severity === ErrorSeverity.CRITICAL) {
      this.notifyAdmin(error, context);
    }
  },
  
  /**
   * Logs an error to the SystemLog sheet
   */
  logToSystemLog(error: Error, context: ErrorContext, severity: ErrorSeverity): void {
    try {
      const logEntry = {
        timestamp: new Date().toISOString(),
        severity,
        module: context.module,
        function: context.function,
        message: error.message,
        stack: error.stack,
        details: context.technicalDetails || {}
      };
      
      // TODO: Implement actual logging to SystemLog sheet
      console.log('Would log to SystemLog:', logEntry);
    } catch (logError) {
      console.error('Failed to log error:', logError);
    }
  },
  
  /**
   * Notifies the user about an error
   */
  notifyUser(message: string, severity: ErrorSeverity): void {
    const ui = SpreadsheetApp.getUi();
    const title = severity === ErrorSeverity.CRITICAL ? 'Critical Error' : 'Error';
    ui.alert(title, message, ui.ButtonSet.OK);
  },
  
  /**
   * Notifies the admin about a critical error
   */
  notifyAdmin(error: Error, _context: ErrorContext): void {
    // TODO: Implement admin notification
    console.log('Would notify admin about:', error.message);
  }
};

/**
 * EventBus for decoupled communication between modules
 */
export const EventBus = {
  events: {} as Record<string, Function[]>,
  
  /**
   * Subscribes to an event
   */
  subscribe(event: string, callback: Function): void {
    if (!this.events[event]) {
      this.events[event] = [];
    }
    this.events[event].push(callback);
  },
  
  /**
   * Publishes an event with optional data
   */
  publish(event: string, data?: any): void {
    if (!this.events[event]) return;
    this.events[event].forEach(callback => {
      try {
        callback(data);
      } catch (error) {
        ErrorHandling.handleError(
          error as Error, 
          { 
            module: 'Core', 
            function: `EventBus.publish(${event})`
          }, 
          ErrorSeverity.ERROR
        );
      }
    });
  },
  
  /**
   * Unsubscribes from an event
   */
  unsubscribe(event: string, callback: Function): void {
    if (!this.events[event]) return;
    this.events[event] = this.events[event].filter(cb => cb !== callback);
  }
};

/**
 * Cache system for performance optimization
 */
export const Cache = {
  data: {} as Record<string, {value: any, expiry: number | null}>,
  
  /**
   * Gets a value from the cache
   */
  get<T>(key: string): T | null {
    const item = this.data[key];
    
    // Check if item exists and not expired
    if (!item) return null;
    
    if (item.expiry !== null && Date.now() > item.expiry) {
      delete this.data[key];
      return null;
    }
    
    return item.value as T;
  },
  
  /**
   * Sets a value in the cache
   */
  set<T>(key: string, value: T, options: CacheOptions = {}): void {
    const expiry = options.expirySeconds 
      ? Date.now() + (options.expirySeconds * 1000)
      : null;
      
    this.data[key] = { value, expiry };
  },
  
  /**
   * Invalidates a cache entry
   */
  invalidate(key: string): void {
    delete this.data[key];
  },
  
  /**
   * Invalidates all cache entries with a given prefix
   */
  invalidateByPrefix(prefix: string): void {
    Object.keys(this.data)
      .filter(key => key.startsWith(prefix))
      .forEach(key => delete this.data[key]);
  },
  
  /**
   * Clears the entire cache
   */
  clear(): void {
    this.data = {};
  }
};

/**
 * FeatureFlags system for feature toggles
 */
export const FeatureFlags = {
  flags: {} as Record<FeatureFlag, boolean>,
  
  /**
   * Initializes feature flags from script properties
   */
  initialize(): void {
    // Default settings
    this.flags = {
      [FeatureFlag.SMART_IMPORT]: true,
      [FeatureFlag.FORM_PROCESSOR]: false,
      [FeatureFlag.ENHANCED_COMMUNICATIONS]: true,
      [FeatureFlag.ADVANCED_REPORTING]: false,
      [FeatureFlag.SESSION_ANALYTICS]: false
    };
    
    // Load from script properties if available
    try {
      const storedFlags = PropertiesService.getScriptProperties().getProperty('FEATURE_FLAGS');
      if (storedFlags) {
        this.flags = JSON.parse(storedFlags);
      }
    } catch (error) {
      console.error('Error loading feature flags:', error);
    }
  },
  
  /**
   * Checks if a feature is enabled
   */
  isEnabled(flag: FeatureFlag): boolean {
    return this.flags[flag] === true;
  },
  
  /**
   * Enables a feature
   */
  enable(flag: FeatureFlag): void {
    this.flags[flag] = true;
    this.saveFlags();
  },
  
  /**
   * Disables a feature
   */
  disable(flag: FeatureFlag): void {
    this.flags[flag] = false;
    this.saveFlags();
  },
  
  /**
   * Saves feature flags to script properties
   */
  saveFlags(): void {
    try {
      PropertiesService.getScriptProperties().setProperty(
        'FEATURE_FLAGS',
        JSON.stringify(this.flags)
      );
    } catch (error) {
      console.error('Error saving feature flags:', error);
    }
  }
};

/**
 * UserManager for role and access management
 */
export const UserManager = {
  /**
   * Gets the current user
   */
  getCurrentUser(): string {
    return Session.getActiveUser().getEmail();
  },
  
  /**
   * Gets the user role
   */
  getUserRole(): UserRole {
    const email = this.getCurrentUser();
    
    // Check script properties for role mappings
    try {
      const roleMap = PropertiesService.getScriptProperties().getProperty('USER_ROLES');
      if (roleMap) {
        const roles = JSON.parse(roleMap);
        if (roles[email]) return roles[email];
      }
    } catch (error) {
      console.error('Error getting user role:', error);
    }
    
    // Default to VIEWER
    return UserRole.VIEWER;
  },
  
  /**
   * Gets the user experience level
   */
  getUserExperienceLevel(): UserExperienceLevel {
    const email = this.getCurrentUser();
    
    // Check script properties for experience levels
    try {
      const levelMap = PropertiesService.getScriptProperties().getProperty('USER_EXPERIENCE_LEVELS');
      if (levelMap) {
        const levels = JSON.parse(levelMap);
        if (levels[email]) return levels[email];
      }
    } catch (error) {
      console.error('Error getting user experience level:', error);
    }
    
    // Default to BEGINNER
    return UserExperienceLevel.BEGINNER;
  },
  
  /**
   * Checks if the user has a required role
   */
  hasPermission(requiredRole: UserRole): boolean {
    const currentRole = this.getUserRole();
    
    // Role hierarchy
    switch (currentRole) {
      case UserRole.ADMIN:
        return true; // Admin can do everything
      case UserRole.COORDINATOR:
        return requiredRole !== UserRole.ADMIN;
      case UserRole.INSTRUCTOR:
        return requiredRole === UserRole.INSTRUCTOR || requiredRole === UserRole.VIEWER;
      case UserRole.VIEWER:
        return requiredRole === UserRole.VIEWER;
      default:
        return false;
    }
  }
};

// Initialize core systems
FeatureFlags.initialize();