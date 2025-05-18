/**
 * Tests for YSLv6Hub Core Module
 */

// Mock Google Apps Script global objects
global.PropertiesService = {
  getScriptProperties: jest.fn().mockReturnValue({
    getProperty: jest.fn().mockImplementation((key) => {
      if (key === 'FEATURE_FLAGS') {
        return JSON.stringify({
          SMART_IMPORT: true,
          FORM_PROCESSOR: false,
          ENHANCED_COMMUNICATIONS: true,
          ADVANCED_REPORTING: false,
          SESSION_ANALYTICS: false
        });
      }
      return null;
    }),
    setProperty: jest.fn()
  })
};

global.Session = {
  getActiveUser: jest.fn().mockReturnValue({
    getEmail: jest.fn().mockReturnValue('test@example.com')
  })
};

global.SpreadsheetApp = {
  getUi: jest.fn().mockReturnValue({
    alert: jest.fn(),
    ButtonSet: { OK: 'OK' }
  })
};

// Import the module to test
import { FeatureFlag, FeatureFlags, UserManager, UserRole, Cache, EventBus } from '../src/01_Core';

describe('FeatureFlags', () => {
  beforeEach(() => {
    // Reset the feature flags before each test
    FeatureFlags.flags = {};
    FeatureFlags.initialize();
  });

  test('should initialize with default values', () => {
    expect(FeatureFlags.isEnabled(FeatureFlag.SMART_IMPORT)).toBe(true);
    expect(FeatureFlags.isEnabled(FeatureFlag.FORM_PROCESSOR)).toBe(false);
    expect(FeatureFlags.isEnabled(FeatureFlag.ENHANCED_COMMUNICATIONS)).toBe(true);
    expect(FeatureFlags.isEnabled(FeatureFlag.ADVANCED_REPORTING)).toBe(false);
    expect(FeatureFlags.isEnabled(FeatureFlag.SESSION_ANALYTICS)).toBe(false);
  });

  test('should enable a feature', () => {
    FeatureFlags.enable(FeatureFlag.FORM_PROCESSOR);
    expect(FeatureFlags.isEnabled(FeatureFlag.FORM_PROCESSOR)).toBe(true);
  });

  test('should disable a feature', () => {
    FeatureFlags.disable(FeatureFlag.SMART_IMPORT);
    expect(FeatureFlags.isEnabled(FeatureFlag.SMART_IMPORT)).toBe(false);
  });

  test('should save flags to script properties when enabled', () => {
    const setPropertyMock = global.PropertiesService.getScriptProperties().setProperty;
    FeatureFlags.enable(FeatureFlag.FORM_PROCESSOR);
    expect(setPropertyMock).toHaveBeenCalledWith('FEATURE_FLAGS', expect.any(String));
    
    // Verify the saved value
    const savedValue = JSON.parse(setPropertyMock.mock.calls[0][1]);
    expect(savedValue[FeatureFlag.FORM_PROCESSOR]).toBe(true);
  });
});

describe('UserManager', () => {
  test('should get current user', () => {
    expect(UserManager.getCurrentUser()).toBe('test@example.com');
  });

  test('should default to VIEWER role', () => {
    expect(UserManager.getUserRole()).toBe(UserRole.VIEWER);
  });

  test('ADMIN should have permission for all roles', () => {
    // Mock getUserRole to return ADMIN
    const originalGetUserRole = UserManager.getUserRole;
    UserManager.getUserRole = jest.fn().mockReturnValue(UserRole.ADMIN);
    
    expect(UserManager.hasPermission(UserRole.ADMIN)).toBe(true);
    expect(UserManager.hasPermission(UserRole.COORDINATOR)).toBe(true);
    expect(UserManager.hasPermission(UserRole.INSTRUCTOR)).toBe(true);
    expect(UserManager.hasPermission(UserRole.VIEWER)).toBe(true);
    
    // Restore original function
    UserManager.getUserRole = originalGetUserRole;
  });

  test('VIEWER should only have permission for VIEWER role', () => {
    // Mock getUserRole to return VIEWER
    const originalGetUserRole = UserManager.getUserRole;
    UserManager.getUserRole = jest.fn().mockReturnValue(UserRole.VIEWER);
    
    expect(UserManager.hasPermission(UserRole.ADMIN)).toBe(false);
    expect(UserManager.hasPermission(UserRole.COORDINATOR)).toBe(false);
    expect(UserManager.hasPermission(UserRole.INSTRUCTOR)).toBe(false);
    expect(UserManager.hasPermission(UserRole.VIEWER)).toBe(true);
    
    // Restore original function
    UserManager.getUserRole = originalGetUserRole;
  });
});

describe('Cache', () => {
  beforeEach(() => {
    // Clear cache before each test
    Cache.clear();
  });

  test('should set and get values', () => {
    Cache.set('testKey', 'testValue');
    expect(Cache.get('testKey')).toBe('testValue');
  });

  test('should return null for non-existent keys', () => {
    expect(Cache.get('nonExistentKey')).toBeNull();
  });

  test('should expire cache entries', () => {
    jest.useFakeTimers();
    const now = Date.now();
    Date.now = jest.fn().mockReturnValue(now);
    
    // Set a cache entry with 5 second expiry
    Cache.set('expiringKey', 'expiringValue', { expirySeconds: 5 });
    
    // Verify it exists
    expect(Cache.get('expiringKey')).toBe('expiringValue');
    
    // Advance time by 6 seconds
    Date.now = jest.fn().mockReturnValue(now + 6000);
    
    // Verify it's expired
    expect(Cache.get('expiringKey')).toBeNull();
    
    // Restore real timers
    jest.useRealTimers();
  });

  test('should invalidate specific keys', () => {
    Cache.set('key1', 'value1');
    Cache.set('key2', 'value2');
    
    Cache.invalidate('key1');
    
    expect(Cache.get('key1')).toBeNull();
    expect(Cache.get('key2')).toBe('value2');
  });

  test('should invalidate by prefix', () => {
    Cache.set('prefix1_key1', 'value1');
    Cache.set('prefix1_key2', 'value2');
    Cache.set('prefix2_key1', 'value3');
    
    Cache.invalidateByPrefix('prefix1_');
    
    expect(Cache.get('prefix1_key1')).toBeNull();
    expect(Cache.get('prefix1_key2')).toBeNull();
    expect(Cache.get('prefix2_key1')).toBe('value3');
  });

  test('should clear the entire cache', () => {
    Cache.set('key1', 'value1');
    Cache.set('key2', 'value2');
    
    Cache.clear();
    
    expect(Cache.get('key1')).toBeNull();
    expect(Cache.get('key2')).toBeNull();
  });
});

describe('EventBus', () => {
  beforeEach(() => {
    // Clear events before each test
    EventBus.events = {};
  });

  test('should publish and subscribe to events', () => {
    const callback = jest.fn();
    EventBus.subscribe('testEvent', callback);
    
    EventBus.publish('testEvent', { data: 'test' });
    
    expect(callback).toHaveBeenCalledWith({ data: 'test' });
  });

  test('should allow multiple subscribers', () => {
    const callback1 = jest.fn();
    const callback2 = jest.fn();
    
    EventBus.subscribe('testEvent', callback1);
    EventBus.subscribe('testEvent', callback2);
    
    EventBus.publish('testEvent', 'test');
    
    expect(callback1).toHaveBeenCalledWith('test');
    expect(callback2).toHaveBeenCalledWith('test');
  });

  test('should allow unsubscribing', () => {
    const callback = jest.fn();
    
    EventBus.subscribe('testEvent', callback);
    EventBus.unsubscribe('testEvent', callback);
    
    EventBus.publish('testEvent');
    
    expect(callback).not.toHaveBeenCalled();
  });

  test('should not fail when publishing to non-existent event', () => {
    expect(() => {
      EventBus.publish('nonExistentEvent');
    }).not.toThrow();
  });

  test('should catch errors in subscribers', () => {
    const errorCallback = jest.fn().mockImplementation(() => {
      throw new Error('Subscriber error');
    });
    
    console.error = jest.fn(); // Mock console.error
    
    EventBus.subscribe('testEvent', errorCallback);
    EventBus.publish('testEvent');
    
    expect(errorCallback).toHaveBeenCalled();
    expect(console.error).toHaveBeenCalled();
  });
});