/**
 * Template Detection Tests
 *
 * Focused tests for template formula detection module (Task Group 5)
 * Tests critical template detection behaviors only
 */

describe('Template Formula Detection', () => {
  // Mock Excel.js types
  interface MockRange {
    formulasR1C1: string[][];
    load: (props: string) => void;
  }

  interface MockContext {
    workbook: {
      getActiveCell: () => MockRange;
      getSelectedRange: () => {
        getCell: (row: number, col: number) => MockRange;
      };
    };
    sync: () => Promise<void>;
  }

  /**
   * Test 1: Active cell template detection returns correct cell
   */
  test('getTemplateCell returns active cell when templateSource is activeCell', async () => {
    const { getTemplateCell } = await import('../src/fillgaps/engine');

    // Mock context
    const mockCell: MockRange = {
      formulasR1C1: [['=R[-1]C']],
      load: jest.fn()
    };

    const mockContext = {
      workbook: {
        getActiveCell: jest.fn(() => mockCell),
        getSelectedRange: jest.fn()
      },
      sync: jest.fn()
    } as unknown as MockContext;

    const result = await getTemplateCell(mockContext as any, 'activeCell');

    expect(mockContext.workbook.getActiveCell).toHaveBeenCalled();
    expect(result).toBe(mockCell);
  });

  /**
   * Test 2: Top-left cell template detection returns correct cell
   */
  test('getTemplateCell returns top-left cell when templateSource is topLeft', async () => {
    const { getTemplateCell } = await import('../src/fillgaps/engine');

    // Mock top-left cell
    const mockCell: MockRange = {
      formulasR1C1: [['=R1C1*2']],
      load: jest.fn()
    };

    const mockRange = {
      getCell: jest.fn(() => mockCell)
    };

    const mockContext = {
      workbook: {
        getActiveCell: jest.fn(),
        getSelectedRange: jest.fn(() => mockRange)
      },
      sync: jest.fn()
    } as unknown as MockContext;

    const result = await getTemplateCell(mockContext as any, 'topLeft');

    expect(mockContext.workbook.getSelectedRange).toHaveBeenCalled();
    expect(mockRange.getCell).toHaveBeenCalledWith(0, 0);
    expect(result).toBe(mockCell);
  });

  /**
   * Test 3: Template validation succeeds with valid formula
   */
  test('validateAndExtractFormula returns R1C1 formula from cell with formula', async () => {
    const { validateAndExtractFormula } = await import('../src/fillgaps/engine');

    const mockCell: MockRange = {
      formulasR1C1: [['=R[-1]C+10']],
      load: jest.fn()
    };

    const mockContext = {
      sync: jest.fn().mockResolvedValue(undefined)
    } as unknown as any;

    const result = await validateAndExtractFormula(mockContext, mockCell as any);

    expect(mockCell.load).toHaveBeenCalledWith('formulasR1C1');
    expect(mockContext.sync).toHaveBeenCalled();
    expect(result).toBe('=R[-1]C+10');
  });

  /**
   * Test 4: Template validation fails when cell has no formula
   */
  test('validateAndExtractFormula throws error when cell has no formula', async () => {
    const { validateAndExtractFormula } = await import('../src/fillgaps/engine');

    // Cell with empty formula (just a value, not a formula)
    const mockCell: MockRange = {
      formulasR1C1: [['']],
      load: jest.fn()
    };

    const mockContext = {
      sync: jest.fn().mockResolvedValue(undefined)
    } as unknown as any;

    await expect(
      validateAndExtractFormula(mockContext, mockCell as any)
    ).rejects.toThrow('Template cell does not contain a formula');
  });

  /**
   * Test 5: Combined template detection returns ITemplateInfo structure
   */
  test('detectTemplateFormula returns template info with cell and formula', async () => {
    // This test will verify the complete template detection flow
    // We'll mock Excel.run to avoid Office.js dependency
    const originalExcel = (global as any).Excel;

    const mockCell = {
      formulasR1C1: [['=SUM(R[-1]C:R[-5]C)']],
      load: jest.fn()
    };

    // Mock Excel.run
    (global as any).Excel = {
      run: jest.fn(async (callback: any) => {
        const mockContext = {
          workbook: {
            getActiveCell: () => mockCell
          },
          sync: jest.fn().mockResolvedValue(undefined)
        };
        return await callback(mockContext);
      })
    };

    const { detectTemplateFormula } = await import('../src/fillgaps/engine');
    const result = await detectTemplateFormula('activeCell');

    expect(result).toHaveProperty('cell');
    expect(result).toHaveProperty('formulaR1C1');
    expect(result.formulaR1C1).toBe('=SUM(R[-1]C:R[-5]C)');
    expect(result.cell).toBe(mockCell);

    // Restore original Excel
    (global as any).Excel = originalExcel;
  });
});
