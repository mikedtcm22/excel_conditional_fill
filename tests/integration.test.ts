/**
 * End-to-End Integration Tests
 *
 * Focused tests for full pipeline integration (Task Group 8)
 * Tests critical end-to-end workflows only
 */

describe('End-to-End Integration', () => {
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

  /**
   * Test 1: User clicks Run with blanks only mode → blanks filled
   */
  test('executeFillOperation fills only blank cells when targetCondition is blanks', async () => {
    const originalExcel = (global as any).Excel;

    // Mock template cell with formula
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=R[-1]C*2']],
      values: [[]],
      formulas: [[]],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Mock selected range with mix of blanks and values
    // Row 0: Template formula
    // Row 1: Blank
    // Row 2: Value (100)
    // Row 3: Blank
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=R[-1]C*2'], [''], [''], ['']],
      values: [[10], [null], [100], [null]],
      formulas: [['=R[-1]C*2'], [''], [''], ['']],
      rowCount: 4,
      columnCount: 1,
      getCell: jest.fn((row: number, _col: number) => {
        // If getCell(0, 0) is called for template source, return template cell
        if (row === 0) {
          return mockTemplateCell;
        }
        // Otherwise return a mock cell for formula writing
        const cell: MockRange = {
          formulasR1C1: [['']],
          values: [[]],
          formulas: [[]],
          rowCount: 1,
          columnCount: 1,
          getCell: jest.fn(),
          load: jest.fn()
        };
        return cell;
      }),
      load: jest.fn()
    };

    // Mock Excel.run - needs to support 3 separate calls (template, identify, write)
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
      templateSource: 'topLeft'
    });

    expect(result.success).toBe(true);
    expect(result.modifiedCount).toBe(2); // Two blanks filled (rows 1 and 3)
    expect(result.error).toBeUndefined();

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 2: User clicks Run with errors only mode → errors filled
   */
  test('executeFillOperation fills only error cells when targetCondition is errors', async () => {
    const originalExcel = (global as any).Excel;

    // Mock template cell with formula
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=VLOOKUP(R1C,Sheet2!R1C1:R10C2,2,FALSE)']],
      values: [[]],
      formulas: [[]],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Mock selected range with mix of errors and values
    // Row 0: Valid formula result
    // Row 1: Error (#N/A)
    // Row 2: Valid value
    // Row 3: Error (#VALUE!)
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=VLOOKUP(R1C,Sheet2!R1C1:R10C2,2,FALSE)'], [''], [''], ['']],
      values: [[100], ['#N/A'], [200], ['#VALUE!']],
      formulas: [['=VLOOKUP(R1C,Sheet2!R1C1:R10C2,2,FALSE)'], [''], [''], ['']],
      rowCount: 4,
      columnCount: 1,
      getCell: jest.fn((_row: number, _col: number) => {
        const cell: MockRange = {
          formulasR1C1: [['']],
          values: [[]],
          formulas: [[]],
          rowCount: 1,
          columnCount: 1,
          getCell: jest.fn(),
          load: jest.fn()
        };
        return cell;
      }),
      load: jest.fn()
    };

    // Mock Excel.run - needs to support 3 separate calls
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
    expect(result.modifiedCount).toBe(2); // Two errors filled (rows 1 and 3)
    expect(result.error).toBeUndefined();

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 3: User clicks Run with blanks + errors mode → both filled
   */
  test('executeFillOperation fills both blanks and errors when targetCondition is both', async () => {
    const originalExcel = (global as any).Excel;

    // Mock template cell with formula
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=R1C1+10']],
      values: [[]],
      formulas: [[]],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Mock selected range with mix of blanks, errors, and values
    // Row 0: Template formula
    // Row 1: Blank
    // Row 2: Error (#DIV/0!)
    // Row 3: Value (100)
    // Row 4: Blank
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=R1C1+10'], [''], [''], [''], ['']],
      values: [[10], [null], ['#DIV/0!'], [100], [null]],
      formulas: [['=R1C1+10'], [''], [''], [''], ['']],
      rowCount: 5,
      columnCount: 1,
      getCell: jest.fn((row: number, _col: number) => {
        // If getCell(0, 0) is called for template source, return template cell
        if (row === 0) {
          return mockTemplateCell;
        }
        // Otherwise return a mock cell for formula writing
        const cell: MockRange = {
          formulasR1C1: [['']],
          values: [[]],
          formulas: [[]],
          rowCount: 1,
          columnCount: 1,
          getCell: jest.fn(),
          load: jest.fn()
        };
        return cell;
      }),
      load: jest.fn()
    };

    // Mock Excel.run - needs to support 3 separate calls
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
      targetCondition: 'both',
      templateSource: 'topLeft'
    });

    expect(result.success).toBe(true);
    expect(result.modifiedCount).toBe(3); // Two blanks + one error filled
    expect(result.error).toBeUndefined();

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 4: No eligible cells found → returns 0 modifications
   */
  test('executeFillOperation returns 0 modifications when no eligible cells found', async () => {
    const originalExcel = (global as any).Excel;

    // Mock template cell with formula
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['=R[-1]C']],
      values: [[]],
      formulas: [[]],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Mock selected range with all values (no blanks or errors)
    const mockSelectedRange: MockRange = {
      formulasR1C1: [['=R[-1]C'], [''], [''], ['']],
      values: [[10], [20], [30], [40]],
      formulas: [['=R[-1]C'], [''], [''], ['']],
      rowCount: 4,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    // Mock Excel.run - needs to support 3 separate calls
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
    expect(result.modifiedCount).toBe(0);
    expect(result.error).toBeUndefined();

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 5: Template has no formula → returns error
   */
  test('executeFillOperation returns error when template cell has no formula', async () => {
    const originalExcel = (global as any).Excel;

    // Mock template cell WITHOUT formula (just a value)
    const mockTemplateCell: MockRange = {
      formulasR1C1: [['']],
      values: [[100]],
      formulas: [['']],
      rowCount: 1,
      columnCount: 1,
      getCell: jest.fn(),
      load: jest.fn()
    };

    const mockSelectedRange: MockRange = {
      formulasR1C1: [[''], [''], ['']],
      values: [[100], [null], [null]],
      formulas: [[''], [''], ['']],
      rowCount: 3,
      columnCount: 1,
      getCell: jest.fn((row: number, _col: number) => {
        // Return template cell for getCell(0, 0)
        if (row === 0) {
          return mockTemplateCell;
        }
        return mockTemplateCell;
      }),
      load: jest.fn()
    };

    // Mock Excel.run - needs to support the template detection call
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
      templateSource: 'topLeft'
    });

    expect(result.success).toBe(false);
    expect(result.modifiedCount).toBe(0);
    expect(result.error).toContain('Template cell does not contain a formula');

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 6: Pipeline handles errors gracefully
   */
  test('executeFillOperation handles errors gracefully and returns error result', async () => {
    const originalExcel = (global as any).Excel;

    // Mock Excel.run to throw error
    (global as any).Excel = {
      run: jest.fn(async (_callback: any) => {
        throw new Error('Excel API error: Network timeout');
      })
    };

    const { executeFillOperation } = await import('../src/fillgaps/engine');

    const result = await executeFillOperation({
      targetCondition: 'blanks',
      templateSource: 'activeCell'
    });

    expect(result.success).toBe(false);
    expect(result.modifiedCount).toBe(0);
    expect(result.error).toBeDefined();
    expect(result.error).toContain('Excel API error');

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });
});
