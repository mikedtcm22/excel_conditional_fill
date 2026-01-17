/**
 * Formula Writing Tests
 *
 * Focused tests for selective formula writing module (Task Group 7)
 * Tests critical formula writing behaviors only
 */

describe('Selective Formula Writing', () => {
  // Mock Excel.js types
  interface MockRange {
    formulasR1C1: string[][];
    getCell: (row: number, col: number) => MockRange;
    load?: (props: string) => void;
  }

  interface MockContext {
    workbook: {
      getSelectedRange: () => MockRange;
    };
    sync: () => Promise<void>;
  }

  /**
   * Test 1: Formula written to eligible cell
   */
  test('writeFormulasToEligibleCells writes formula to single eligible cell', async () => {
    const originalExcel = (global as any).Excel;

    const mockCell: MockRange = {
      formulasR1C1: [['']],
      getCell: jest.fn()
    };

    const mockRange: MockRange = {
      formulasR1C1: [[]],
      getCell: jest.fn(() => mockCell)
    };

    // Mock Excel.run
    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getSelectedRange: () => mockRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { writeFormulasToEligibleCells } = await import('../src/fillgaps/engine');

    const eligibleCells = [{ row: 0, col: 0 }];
    const templateFormula = '=R[-1]C';

    const result = await writeFormulasToEligibleCells(eligibleCells, templateFormula);

    expect(mockRange.getCell).toHaveBeenCalledWith(0, 0);
    expect(mockCell.formulasR1C1).toEqual([['=R[-1]C']]);
    expect(result).toBe(1);

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 2: Formulas written to multiple eligible cells
   */
  test('writeFormulasToEligibleCells writes formula to multiple eligible cells', async () => {
    const originalExcel = (global as any).Excel;

    const mockCells: Map<string, MockRange> = new Map();

    // Create mock cells for each coordinate
    const createMockCell = () => ({
      formulasR1C1: [['']],
      getCell: jest.fn()
    });

    const mockRange: MockRange = {
      formulasR1C1: [[]],
      getCell: jest.fn((row: number, col: number) => {
        const key = `${row},${col}`;
        if (!mockCells.has(key)) {
          mockCells.set(key, createMockCell());
        }
        return mockCells.get(key)!;
      })
    };

    // Mock Excel.run
    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getSelectedRange: () => mockRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { writeFormulasToEligibleCells } = await import('../src/fillgaps/engine');

    const eligibleCells = [
      { row: 0, col: 0 },
      { row: 2, col: 0 },
      { row: 4, col: 0 }
    ];
    const templateFormula = '=R[-1]C*2';

    const result = await writeFormulasToEligibleCells(eligibleCells, templateFormula);

    expect(mockRange.getCell).toHaveBeenCalledTimes(3);
    expect(mockRange.getCell).toHaveBeenCalledWith(0, 0);
    expect(mockRange.getCell).toHaveBeenCalledWith(2, 0);
    expect(mockRange.getCell).toHaveBeenCalledWith(4, 0);
    expect(result).toBe(3);

    // Verify each cell received the formula
    mockCells.forEach((cell) => {
      expect(cell.formulasR1C1).toEqual([['=R[-1]C*2']]);
    });

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 3: Empty eligible cells array returns 0 without calling Excel.run
   */
  test('writeFormulasToEligibleCells returns 0 when no eligible cells', async () => {
    const originalExcel = (global as any).Excel;

    // Mock Excel.run to track if it's called
    const excelRunSpy = jest.fn();
    (global as any).Excel = {
      run: excelRunSpy
    };

    const { writeFormulasToEligibleCells } = await import('../src/fillgaps/engine');

    const eligibleCells: Array<{ row: number; col: number }> = [];
    const templateFormula = '=R[-1]C';

    const result = await writeFormulasToEligibleCells(eligibleCells, templateFormula);

    // Should return 0 immediately without calling Excel.run
    expect(result).toBe(0);
    expect(excelRunSpy).not.toHaveBeenCalled();

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 4: context.sync called once after all formulas set
   */
  test('writeFormulasToEligibleCells calls context.sync once after setting all formulas', async () => {
    const originalExcel = (global as any).Excel;

    const mockCell: MockRange = {
      formulasR1C1: [['']],
      getCell: jest.fn()
    };

    const mockRange: MockRange = {
      formulasR1C1: [[]],
      getCell: jest.fn(() => mockCell)
    };

    const syncSpy = jest.fn().mockResolvedValue(undefined);

    // Mock Excel.run
    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getSelectedRange: () => mockRange
          },
          sync: syncSpy
        };
        return await callback(mockContext);
      })
    };

    const { writeFormulasToEligibleCells } = await import('../src/fillgaps/engine');

    const eligibleCells = [
      { row: 0, col: 0 },
      { row: 1, col: 0 },
      { row: 2, col: 0 }
    ];
    const templateFormula = '=SUM(R[-1]C:R[-3]C)';

    await writeFormulasToEligibleCells(eligibleCells, templateFormula);

    // context.sync should be called exactly once
    expect(syncSpy).toHaveBeenCalledTimes(1);

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });

  /**
   * Test 5: Formula assignment uses correct R1C1 format
   */
  test('writeFormulasToEligibleCells assigns formula in correct R1C1 2D array format', async () => {
    const originalExcel = (global as any).Excel;

    const mockCell: MockRange = {
      formulasR1C1: [['']],
      getCell: jest.fn()
    };

    const mockRange: MockRange = {
      formulasR1C1: [[]],
      getCell: jest.fn(() => mockCell)
    };

    // Mock Excel.run
    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext: MockContext = {
          workbook: {
            getSelectedRange: () => mockRange
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { writeFormulasToEligibleCells } = await import('../src/fillgaps/engine');

    const eligibleCells = [{ row: 1, col: 2 }];
    const templateFormula = '=R1C1';

    await writeFormulasToEligibleCells(eligibleCells, templateFormula);

    // Verify formula was assigned as 2D array: [[formula]]
    expect(mockCell.formulasR1C1).toEqual([['=R1C1']]);

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });
});
