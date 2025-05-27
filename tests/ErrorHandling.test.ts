/**
 * Tests for Error Handling Module
 */

// Import the module which sets up globals
import '../02_ErrorHandling';

// Use the global ErrorHandling that was created
declare const ErrorHandling: any;

describe('ErrorHandling Module', () => {
  beforeEach(() => {
    // Clear all mocks before each test
    jest.clearAllMocks();
    
    // Reset properties mock
    const mockProperties = PropertiesService.getScriptProperties();
    (mockProperties.getProperty as jest.Mock).mockReturnValue(null);
  });

  describe('initializeErrorHandling', () => {
    it('should initialize successfully with default log level', () => {
      const mockSheet = {
        getName: jest.fn(() => 'SystemLogs'),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          setFontColors: jest.fn(),
          setHorizontalAlignment: jest.fn(),
          setVerticalAlignment: jest.fn(),
          setWrap: jest.fn(),
        })),
      };
      
      const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      (mockSpreadsheet.getSheetByName as jest.Mock).mockReturnValue(mockSheet);
      
      const result = ErrorHandling.initializeErrorHandling();
      
      expect(result).toBe(true);
      expect(mockSpreadsheet.getSheetByName).toHaveBeenCalledWith('SystemLogs');
    });

    it('should use stored log level from properties', () => {
      const mockProperties = PropertiesService.getScriptProperties();
      (mockProperties.getProperty as jest.Mock).mockReturnValue('DEBUG');
      
      const mockSheet = {
        getName: jest.fn(() => 'SystemLogs'),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          setFontColors: jest.fn(),
          setHorizontalAlignment: jest.fn(),
          setVerticalAlignment: jest.fn(),
          setWrap: jest.fn(),
        })),
      };
      
      const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      (mockSpreadsheet.getSheetByName as jest.Mock).mockReturnValue(mockSheet);
      
      ErrorHandling.initializeErrorHandling();
      
      expect(mockProperties.getProperty).toHaveBeenCalledWith('logLevel');
    });

    it('should handle initialization errors gracefully', () => {
      const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      (mockSpreadsheet.getSheetByName as jest.Mock).mockImplementation(() => {
        throw new Error('Sheet access error');
      });
      
      const result = ErrorHandling.initializeErrorHandling();
      
      expect(result).toBe(false);
      expect(Logger.log).toHaveBeenCalled();
    });
  });

  describe('logMessage', () => {
    beforeEach(() => {
      // Initialize system first
      const mockSheet = {
        getName: jest.fn(() => 'SystemLogs'),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          setFontColors: jest.fn(),
          setHorizontalAlignment: jest.fn(),
          setVerticalAlignment: jest.fn(),
          setWrap: jest.fn(),
        })),
        getLastRow: jest.fn(() => 1),
        insertRowAfter: jest.fn(),
      };
      
      const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      (mockSpreadsheet.getSheetByName as jest.Mock).mockReturnValue(mockSheet);
      
      ErrorHandling.initializeErrorHandling();
    });

    it('should log messages with proper format', () => {
      const mockSheet = {
        getLastRow: jest.fn(() => 1),
        insertRowAfter: jest.fn(),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          setFontColors: jest.fn(),
          setHorizontalAlignment: jest.fn(),
          setVerticalAlignment: jest.fn(),
          setWrap: jest.fn(),
        })),
      };
      
      const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      (mockSpreadsheet.getSheetByName as jest.Mock).mockReturnValue(mockSheet);
      
      ErrorHandling.logMessage('Test message', 'INFO', 'TestFunction');
      
      expect(mockSheet.insertRowAfter).toHaveBeenCalledWith(1);
      expect(mockSheet.getRange).toHaveBeenCalled();
    });

    it('should respect log level filtering', () => {
      const mockSheet = {
        getLastRow: jest.fn(() => 1),
        insertRowAfter: jest.fn(),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          setFontColors: jest.fn(),
          setHorizontalAlignment: jest.fn(),
          setVerticalAlignment: jest.fn(),
          setWrap: jest.fn(),
        })),
      };
      
      const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      (mockSpreadsheet.getSheetByName as jest.Mock).mockReturnValue(mockSheet);
      
      // Set log level to ERROR
      ErrorHandling.setLogLevel('ERROR');
      
      // Try to log INFO message (should be filtered)
      ErrorHandling.logMessage('Info message', 'INFO', 'TestFunction');
      
      // INFO should not be logged when level is ERROR
      expect(mockSheet.insertRowAfter).not.toHaveBeenCalled();
      
      // ERROR should be logged
      ErrorHandling.logMessage('Error message', 'ERROR', 'TestFunction');
      expect(mockSheet.insertRowAfter).toHaveBeenCalled();
    });
  });

  describe('handleError', () => {
    it('should show user-friendly error messages', () => {
      const mockUi = SpreadsheetApp.getUi();
      const mockError = new Error('Test error');
      
      ErrorHandling.handleError(mockError, 'TestFunction', 'An error occurred in TestFunction', true);
      
      expect(mockUi.alert).toHaveBeenCalledWith(
        'Operation Error',
        expect.stringContaining('An error occurred in TestFunction'),
        expect.anything()
      );
    });

    it('should not show UI alert when showAlert is false', () => {
      const mockUi = SpreadsheetApp.getUi();
      const mockError = new Error('Test error');
      
      ErrorHandling.handleError(mockError, 'TestFunction', 'An error occurred in TestFunction', false);
      
      expect(mockUi.alert).not.toHaveBeenCalled();
    });
  });

  describe('clearLog', () => {
    it('should clear all logs except header when user confirms', () => {
      const mockSheet = {
        getLastRow: jest.fn(() => 10),
        deleteRows: jest.fn(),
        insertRowAfter: jest.fn(),
        getRange: jest.fn(() => ({
          setValues: jest.fn(),
          setFontColors: jest.fn(),
          setHorizontalAlignment: jest.fn(),
          setVerticalAlignment: jest.fn(),
          setWrap: jest.fn(),
        })),
      };
      
      const mockUi = SpreadsheetApp.getUi();
      const mockSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      (mockSpreadsheet.getSheetByName as jest.Mock).mockReturnValue(mockSheet);
      (mockUi.alert as jest.Mock)
        .mockReturnValueOnce(mockUi.Button.YES)  // For confirmation dialog
        .mockReturnValue(mockUi.Button.OK);      // For success message
      
      ErrorHandling.clearLog();
      
      expect(mockSheet.deleteRows).toHaveBeenCalledWith(2, 9);
      expect(mockUi.alert).toHaveBeenCalledWith(
        'Clear System Log',
        expect.stringContaining('Are you sure'),
        mockUi.ButtonSet.YES_NO
      );
    });
  });
});