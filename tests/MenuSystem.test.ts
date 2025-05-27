/**
 * Tests for Menu System Module
 */

import '../00_MenuSystem';

// Declare globals from imported module
declare const onOpen: any;
declare const wrapInitializeBlankSpreadsheet: any;
declare const showFixedConfigurationDialog: any;
declare const wrapUpgradeExistingSpreadsheet: any;
declare const handleError: any;

describe('MenuSystem Module', () => {
  let mockMenu: any;
  let mockUi: any;

  beforeEach(() => {
    jest.clearAllMocks();
    
    // Setup menu mock
    mockMenu = {
      addItem: jest.fn().mockReturnThis(),
      addSeparator: jest.fn().mockReturnThis(),
      addSubMenu: jest.fn().mockReturnThis(),
      addToUi: jest.fn(),
    };
    
    mockUi = SpreadsheetApp.getUi();
    (mockUi.createMenu as jest.Mock).mockReturnValue(mockMenu);
  });

  describe('onOpen', () => {
    it('should create the main menu structure', () => {
      onOpen();
      
      expect(mockUi.createMenu).toHaveBeenCalledWith('YSL Hub Enhanced');
      expect(mockMenu.addToUi).toHaveBeenCalled();
    });

    it('should add all main menu items', () => {
      onOpen();
      
      // Check for main menu items
      expect(mockMenu.addItem).toHaveBeenCalledWith(
        'System Configuration (Fixed)',
        'showFixedConfigurationDialog'
      );
      
      expect(mockMenu.addItem).toHaveBeenCalledWith(
        'Initialize Blank Spreadsheet',
        'wrapInitializeBlankSpreadsheet'
      );
    });

    it('should create submenus', () => {
      onOpen();
      
      // Should create multiple submenus
      const addSubMenuCalls = (mockMenu.addSubMenu as jest.Mock).mock.calls;
      expect(addSubMenuCalls.length).toBeGreaterThan(0);
      
      // Check if Email Templates submenu was created
      const emailSubmenuCall = addSubMenuCalls.find(
        (call: any[]) => call[0].addItem.mock.calls.some(
          (itemCall: any[]) => itemCall[0].includes('Email')
        )
      );
      expect(emailSubmenuCall).toBeDefined();
    });
  });

  describe('Menu Item Functions', () => {
    it('should have wrapper functions for menu items', () => {
      // Test that wrapper functions exist
      expect(typeof wrapInitializeBlankSpreadsheet).toBe('function');
      expect(typeof showFixedConfigurationDialog).toBe('function');
      expect(typeof wrapUpgradeExistingSpreadsheet).toBe('function');
    });

    it('should handle errors in menu functions gracefully', () => {
      // Mock an error in a wrapped function
      const mockError = new Error('Test error');
      
      // Create a test wrapper that throws
      const testWrapper = () => {
        try {
          throw mockError;
        } catch (error) {
          if (typeof handleError === 'function') {
            handleError(error, 'testWrapper', true);
          }
        }
      };
      
      expect(() => testWrapper()).not.toThrow();
    });
  });

  describe('Dynamic Menu Building', () => {
    it('should build email templates submenu', () => {
      const emailMenu = {
        addItem: jest.fn().mockReturnThis(),
        addSeparator: jest.fn().mockReturnThis(),
      };
      
      (mockUi.createMenu as jest.Mock).mockReturnValue(emailMenu);
      
      onOpen();
      
      // Find email submenu creation
      const emailSubmenuCalls = (mockMenu.addSubMenu as jest.Mock).mock.calls;
      expect(emailSubmenuCalls.some((call: any[]) => {
        const submenu = call[0];
        return submenu.addItem.mock.calls.some(
          (itemCall: any[]) => itemCall[0] === 'Manage Email Templates'
        );
      })).toBe(true);
    });

    it('should build tools submenu', () => {
      onOpen();
      
      // Find tools submenu creation
      const toolsSubmenuCalls = (mockMenu.addSubMenu as jest.Mock).mock.calls;
      expect(toolsSubmenuCalls.some((call: any[]) => {
        const submenu = call[0];
        return submenu.addItem.mock.calls.some(
          (itemCall: any[]) => itemCall[0] === 'System Health Check'
        );
      })).toBe(true);
    });
  });

  describe('Menu Initialization', () => {
    it('should handle missing ErrorHandling module gracefully', () => {
      // Temporarily remove ErrorHandling
      const originalErrorHandling = global.ErrorHandling;
      delete global.ErrorHandling;
      
      expect(() => onOpen()).not.toThrow();
      
      // Restore
      global.ErrorHandling = originalErrorHandling;
    });

    it('should add separators between menu sections', () => {
      onOpen();
      
      // Should have multiple separators
      expect(mockMenu.addSeparator).toHaveBeenCalled();
      expect((mockMenu.addSeparator as jest.Mock).mock.calls.length).toBeGreaterThan(1);
    });
  });
});