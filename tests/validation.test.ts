/**
 * Validation Logic Tests
 *
 * Focused tests for preflight validation and error display (Task Group 10)
 * Tests validation of template formulas and error dialog display
 */

// Mock Excel.js API
const mockSync = jest.fn().mockResolvedValue(undefined);
const mockLoad = jest.fn();

interface MockRange {
  formulasR1C1: string[][];
  load: jest.Mock;
}

interface MockContext {
  workbook: {
    getActiveCell: () => MockRange;
    getSelectedRange: () => MockRange;
  };
  sync: jest.Mock;
}

// Store the current mock template cell for dynamic testing
let mockTemplateCell: MockRange;

// Mock Excel global
(global as any).Excel = {
  run: jest.fn(async (callback: (context: MockContext) => Promise<any>) => {
    const mockContext: MockContext = {
      workbook: {
        getActiveCell: () => mockTemplateCell,
        getSelectedRange: () => ({
          formulasR1C1: [['']],
          load: mockLoad,
          getCell: () => mockTemplateCell
        }) as any
      },
      sync: mockSync
    };
    return await callback(mockContext);
  })
};

describe('Preflight Validation', () => {
  beforeEach(() => {
    jest.clearAllMocks();
  });

  /**
   * Test 1: validatePreflightConditions returns error when active cell has no formula
   */
  test('validatePreflightConditions returns error when active cell has no formula', async () => {
    // Setup mock template cell WITHOUT formula (just a value)
    mockTemplateCell = {
      formulasR1C1: [['']],
      load: mockLoad
    };

    const { validatePreflightConditions } = await import('../src/fillgaps/validation');

    const result = await validatePreflightConditions('activeCell');

    expect(result.valid).toBe(false);
    expect(result.error).toBe('Active cell must contain a formula');
  });

  /**
   * Test 2: validatePreflightConditions returns success when active cell has formula
   */
  test('validatePreflightConditions returns success when active cell has formula', async () => {
    // Setup mock template cell WITH formula
    mockTemplateCell = {
      formulasR1C1: [['=R[-1]C*2']],
      load: mockLoad
    };

    const { validatePreflightConditions } = await import('../src/fillgaps/validation');

    const result = await validatePreflightConditions('activeCell');

    expect(result.valid).toBe(true);
    expect(result.error).toBeUndefined();
  });

  /**
   * Test 3: showErrorDialog displays message (mock alert)
   */
  test('showErrorDialog displays message using alert', () => {
    // Mock global alert
    const originalAlert = global.alert;
    global.alert = jest.fn();

    // Import synchronously since showErrorDialog is not async
    const { showErrorDialog } = require('../src/fillgaps/validation');

    showErrorDialog('Test error message');

    expect(global.alert).toHaveBeenCalledTimes(1);
    expect(global.alert).toHaveBeenCalledWith('Test error message');

    // Restore alert
    global.alert = originalAlert;
  });

  /**
   * Test 4: validatePreflightConditions works with topLeft template source
   */
  test('validatePreflightConditions returns error when topLeft cell has no formula', async () => {
    // Setup mock template cell WITHOUT formula
    mockTemplateCell = {
      formulasR1C1: [['']],
      load: mockLoad
    };

    // Override Excel.run to handle topLeft case
    (global as any).Excel.run = jest.fn(async (callback: (context: any) => Promise<any>) => {
      const mockSelectedRange = {
        formulasR1C1: [['']],
        load: mockLoad,
        getCell: () => mockTemplateCell
      };

      const mockContext = {
        workbook: {
          getActiveCell: () => mockTemplateCell,
          getSelectedRange: () => mockSelectedRange
        },
        sync: mockSync
      };
      return await callback(mockContext);
    });

    const { validatePreflightConditions } = await import('../src/fillgaps/validation');

    const result = await validatePreflightConditions('topLeft');

    expect(result.valid).toBe(false);
    expect(result.error).toBe('Active cell must contain a formula');
  });
});
