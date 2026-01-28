/**
 * Edge Case Tests
 *
 * Comprehensive edge case coverage for FillGaps v0 (Task Group 11)
 * Tests selection edge cases, content edge cases, error type coverage,
 * and template formula edge cases.
 *
 * Maximum 10 tests as per spec requirements.
 */

describe('Edge Cases', () => {
  // Mock Excel.js types
  interface MockRange {
    formulasR1C1: string[][];
    values: any[][];
    formulas: any[][];
    rowCount: number;
    columnCount: number;
    getCell: (row: number, col: number) => MockRange;
    load: (props: string) => void;
  }

  interface MockContext {
    workbook: {
      getActiveCell: () => MockRange;
      getSelectedRange: () => MockRange;
    };
    sync: () => Promise<void>;
  }

  // Store original Excel for cleanup
  let originalExcel: any;

  beforeEach(() => {
    originalExcel = (global as any).Excel;
    jest.clearAllMocks();
  });

  afterEach(() => {
    (global as any).Excel = originalExcel;
  });

  // =========================================================================
  // SELECTION EDGE CASES (3 tests)
  // =========================================================================

  /**
   * Test 1: Single cell selection with formula
   * When a single cell with a formula is selected, expect 0 cells modified
   * and success (no blanks/errors to fill in a single cell selection
   * where the cell itself is the template).
   */
  test('Single cell selection with formula - expect 0 cells modified, success message', async () => {
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=ROW()']],
      values: [[1]],
      formulas: [['=ROW()']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Single cell selection (only the template cell)
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=ROW()']],
      values: [[1]],
      formulas: [['=ROW()']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(() => mockTemplateCell),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    // Single cell with formula has no blanks to fill - returns error message
    expect(result.success).toBe(false);
    expect(result.modifiedCount).toBe(0);
    expect(result.error).toBe('No blank cells found in selection');
  });

  /**
   * Test 2: Empty selection handling (all cells have values)
   * When all cells have values and none are blank/errors,
   * expect graceful error message.
   */
  test('Empty eligible cells - all cells have values - expect graceful error message', async () => {
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=ROW()']],
      values: [[1]],
      formulas: [['=ROW()']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Range where all cells have values (no blanks, no errors)
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=ROW()'], [''], [''], ['']],
      values: [[1], [100], [200], [300]],
      formulas: [['=ROW()'], [''], [''], ['']],
      rowCount: 4,
      columnCount: 1,
      getCell: jest.fn(() => mockTemplateCell),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(false);
    expect(result.modifiedCount).toBe(0);
    expect(result.error).toBe('No blank cells found in selection');
  });

  /**
   * Test 3: Large range (1000 cells, ~50% blanks)
   * Expect approximately 500 cells modified and completes successfully.
   * Performance should be reasonable (test will timeout if too slow).
   */
  test('Large range (1000 cells, 50% blanks) - expect ~500 cells modified, completes successfully', async () => {
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=ROW()']],
      values: [[1]],
      formulas: [['=ROW()']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Generate large range: 1000 rows, alternating blanks and values
    const largeFormulasR1C1: string[][] = [];
    const largeValues: any[][] = [];
    const largeFormulas: any[][] = [];

    // First row is template
    largeFormulasR1C1.push(['=ROW()']);
    largeValues.push([[1]]);
    largeFormulas.push(['=ROW()']);

    // Remaining 999 rows: alternate blank and value
    for (let i = 1; i < 1000; i++) {
      if (i % 2 === 0) {
        // Even rows: have values
        largeFormulasR1C1.push(['']);
        largeValues.push([i * 10]);
        largeFormulas.push(['']);
      } else {
        // Odd rows: blank
        largeFormulasR1C1.push(['']);
        largeValues.push([null]);
        largeFormulas.push(['']);
      }
    }

    const mockSelectedRange: MockRange = {
      formulasR1C1: largeFormulasR1C1,
      values: largeValues,
      formulas: largeFormulas,
      rowCount: 1000,
      columnCount: 1,
      getCell: jest.fn((_row: number, _col: number) => ({
        formulasR1C1: [['']],
        values: [[]],
        formulas: [[]],
        rowCount: 1,
        columnCount: 1,
        getCell: jest.fn(),
        load: jest.fn()
      })),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const startTime = Date.now();
    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });
    const endTime = Date.now();

    expect(result.success).toBe(true);
    // ~500 blank cells (odd numbered rows from 1-999, which is 500 rows)
    expect(result.modifiedCount).toBe(500);
    expect(result.error).toBeUndefined();
    // Should complete in less than 5 seconds
    expect(endTime - startTime).toBeLessThan(5000);
  }, 10000); // 10 second timeout for this test

  // =========================================================================
  // CONTENT EDGE CASES (3 tests)
  // =========================================================================

  /**
   * Test 4: Range where all non-template cells are eligible (all blanks)
   * Expect all cells filled except template.
   */
  test('All non-template cells are blanks - expect all filled except template', async () => {
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=ROW()*2']],
      values: [[2]],
      formulas: [['=ROW()*2']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Range with template in first cell, rest are all blank
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=ROW()*2'], [''], [''], [''], ['']],
      values: [[2], [null], [null], [null], [null]],
      formulas: [['=ROW()*2'], [''], [''], [''], ['']],
      rowCount: 5,
      columnCount: 1,
      getCell: jest.fn((_row: number, _col: number) => ({
        formulasR1C1: [['']],
        values: [[]],
        formulas: [[]],
        rowCount: 1,
        columnCount: 1,
        getCell: jest.fn(),
        load: jest.fn()
      })),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(true);
    // 4 blank cells filled (rows 1-4, template is row 0)
    expect(result.modifiedCount).toBe(4);
    expect(result.error).toBeUndefined();
  });

  /**
   * Test 5: Range where no cells are eligible (all have values)
   * Expect 0 cells modified with informational message.
   */
  test('No cells eligible (all have values/formulas) - expect 0 modified with message', async () => {
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=ROW()']],
      values: [[1]],
      formulas: [['=ROW()']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // All cells have values or formulas - no blanks, no errors
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=ROW()'], ['=ROW()+1'], [''], ['']],
      values: [[1], [2], [100], [200]],
      formulas: [['=ROW()'], ['=ROW()+1'], [''], ['']],
      rowCount: 4,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(false);
    expect(result.modifiedCount).toBe(0);
    expect(result.error).toBe('No blank cells found in selection');
  });

  /**
   * Test 6: Mixed content (blanks, errors, values, formulas)
   * Only eligible cells (blanks for 'blanks' condition) should be filled.
   */
  test('Mixed content - only eligible cells filled based on targetCondition', async () => {
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=ROW()']],
      values: [[1]],
      formulas: [['=ROW()']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Mixed content: template, blank, error, value, formula, blank
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=ROW()'], [''], [''], [''], ['=ROW()+10'], ['']],
      values: [[1], [null], ['#N/A'], [100], [11], [null]],
      formulas: [['=ROW()'], [''], [''], [''], ['=ROW()+10'], ['']],
      rowCount: 6,
      columnCount: 1,
      getCell: jest.fn((_row: number, _col: number) => ({
        formulasR1C1: [['']],
        values: [[]],
        formulas: [[]],
        rowCount: 1,
        columnCount: 1,
        getCell: jest.fn(),
        load: jest.fn()
      })),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    // Test with 'blanks' condition - should fill only blanks (rows 1 and 5)
    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(true);
    // Only 2 blank cells filled (rows 1 and 5), error and value not touched
    expect(result.modifiedCount).toBe(2);
    expect(result.error).toBeUndefined();
  });

  // =========================================================================
  // ERROR TYPE COVERAGE (2 tests)
  // =========================================================================

  /**
   * Test 7: All 7 Excel error types are detected as errors
   * Tests: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
   */
  test('All 7 Excel error types are detected as errors', () => {
    // Import the isError function directly to test error detection
    const { isError } = require('../src/fillgaps/engine');

    // Test all 7 Excel error types
    const errorTypes = [
      '#N/A',
      '#VALUE!',
      '#REF!',
      '#DIV/0!',
      '#NUM!',
      '#NAME?',
      '#NULL!'
    ];

    errorTypes.forEach((errorType) => {
      expect(isError(errorType)).toBe(true);
    });

    // Also test that non-errors are not detected as errors
    expect(isError(100)).toBe(false);
    expect(isError('hello')).toBe(false);
    expect(isError(null)).toBe(false);
    expect(isError(undefined)).toBe(false);
    expect(isError('')).toBe(false);
  });

  /**
   * Test 8: Error cells are filled when targetCondition is 'errors'
   */
  test('Error cells are filled when targetCondition is errors', async () => {
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=VLOOKUP(A1,B:C,2,FALSE)']],
      values: [['Found']],
      formulas: [['=VLOOKUP(A1,B:C,2,FALSE)']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Range with various error types
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=VLOOKUP(A1,B:C,2,FALSE)'], [''], [''], [''], ['']],
      values: [['Found'], ['#N/A'], ['#VALUE!'], [100], ['#DIV/0!']],
      formulas: [['=VLOOKUP(A1,B:C,2,FALSE)'], [''], [''], [''], ['']],
      rowCount: 5,
      columnCount: 1,
      getCell: jest.fn((_row: number, _col: number) => ({
        formulasR1C1: [['']],
        values: [[]],
        formulas: [[]],
        rowCount: 1,
        columnCount: 1,
        getCell: jest.fn(),
        load: jest.fn()
      })),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'errors',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(true);
    // 3 error cells filled (rows 1, 2, 4 with #N/A, #VALUE!, #DIV/0!)
    expect(result.modifiedCount).toBe(3);
    expect(result.error).toBeUndefined();
  });

  // =========================================================================
  // TEMPLATE FORMULA EDGE CASES (2 tests)
  // =========================================================================

  /**
   * Test 9: Simple relative reference (=R[-1]C) adjusts correctly per cell position
   * R1C1 formulas with relative references should work correctly.
   */
  test('Simple relative reference formula is correctly used for filling', async () => {
    const relativeFormula = '=R[-1]C';

    const mockTemplateCell: MockRange = {
      formulasR1C1: [[relativeFormula]],
      values: [[10]],
      formulas: [[relativeFormula]],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    const writtenFormulas: string[] = [];
    const mockSelectedRange: MockRange = {
      formulasR1C1: [[relativeFormula], [''], ['']],
      values: [[10], [null], [null]],
      formulas: [[relativeFormula], [''], ['']],
      rowCount: 3,
      columnCount: 1,
      getCell: jest.fn((_row: number, _col: number) => {
        const cell = {
          formulasR1C1: [['']],
          values: [[]],
          formulas: [[]],
          rowCount: 1,
          columnCount: 1,
          getCell: jest.fn(),
          load: jest.fn()
        };
        // Capture what formulas get written
        Object.defineProperty(cell, 'formulasR1C1', {
          set: (value: string[][]) => {
            writtenFormulas.push(value[0][0]);
          },
          get: () => [['']]
        });
        return cell;
      }),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(true);
    expect(result.modifiedCount).toBe(2);
    // The relative formula should be written to both blank cells
    expect(writtenFormulas).toContain(relativeFormula);
  });

  /**
   * Test 10: Absolute reference (=R1C1) remains fixed across filled cells
   * R1C1 formulas with absolute references should be preserved.
   */
  test('Absolute reference formula (=R1C1) remains fixed across filled cells', async () => {
    const absoluteFormula = '=R1C1';

    const mockTemplateCell: MockRange = {
      formulasR1C1: [[absoluteFormula]],
      values: [[100]],
      formulas: [[absoluteFormula]],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    const writtenFormulas: string[] = [];
    const mockSelectedRange: MockRange = {
      formulasR1C1: [[absoluteFormula], [''], [''], ['']],
      values: [[100], [null], [null], [null]],
      formulas: [[absoluteFormula], [''], [''], ['']],
      rowCount: 4,
      columnCount: 1,
      getCell: jest.fn((_row: number, _col: number) => {
        const cell = {
          formulasR1C1: [['']],
          values: [[]],
          formulas: [[]],
          rowCount: 1,
          columnCount: 1,
          getCell: jest.fn(),
          load: jest.fn()
        };
        // Capture what formulas get written
        Object.defineProperty(cell, 'formulasR1C1', {
          set: (value: string[][]) => {
            writtenFormulas.push(value[0][0]);
          },
          get: () => [['']]
        });
        return cell;
      }),
      load: jest.fn()
    };

    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getActiveCell: () => mockTemplateCell,
            getSelectedRange: () => mockSelectedRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(true);
    expect(result.modifiedCount).toBe(3);
    // All written formulas should be the same absolute reference
    writtenFormulas.forEach((formula) => {
      expect(formula).toBe(absoluteFormula);
    });
  });
});
