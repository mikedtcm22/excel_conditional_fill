/// <reference types="office-js" />

/**
 * FillGaps Core Engine
 *
 * Core logic for template formula detection, cell eligibility identification,
 * and selective formula writing for the FillGaps Excel add-in.
 */

import { ITemplateInfo, IFillOptions, IFillResult } from './types';
import { validatePreflightConditions, getNoEligibleCellsMessage } from './validation';

/**
 * Gets the template cell based on the user's selection
 *
 * @param context - Excel.RequestContext for API calls
 * @param templateSource - Source of template: "activeCell" or "topLeft"
 * @returns Promise resolving to the template cell range
 */
export async function getTemplateCell(
  context: Excel.RequestContext,
  templateSource: string
): Promise<Excel.Range> {
  let templateCell: Excel.Range;

  if (templateSource === 'activeCell') {
    // Use the currently active cell
    templateCell = context.workbook.getActiveCell();
  } else if (templateSource === 'topLeft') {
    // Use the top-left cell of the selection
    const selectedRange = context.workbook.getSelectedRange();
    templateCell = selectedRange.getCell(0, 0);
  } else {
    throw new Error(`Invalid template source: ${templateSource}`);
  }

  // Load cell properties for later use
  templateCell.load('formulasR1C1');

  return templateCell;
}

/**
 * Validates that the template cell contains a formula and extracts it
 *
 * @param context - Excel.RequestContext for API calls
 * @param templateCell - The template cell range to validate
 * @returns Promise resolving to the R1C1 formula string
 * @throws Error if the template cell does not contain a formula
 */
export async function validateAndExtractFormula(
  context: Excel.RequestContext,
  templateCell: Excel.Range
): Promise<string> {
  // Load the formulasR1C1 property
  templateCell.load('formulasR1C1');

  // Sync to get the actual values
  await context.sync();

  // Extract the formula from the 2D array
  const formula = templateCell.formulasR1C1[0][0];

  // Validate that a formula exists
  if (!formula || formula === '') {
    throw new Error('Active cell must contain a formula');
  }

  return formula;
}

/**
 * Detects and validates the template formula based on user selection
 *
 * This is the main entry point for template detection that combines
 * cell identification and formula validation.
 *
 * @param templateSource - Source of template: "activeCell" or "topLeft"
 * @returns Promise resolving to template information (cell and formula)
 */
export async function detectTemplateFormula(
  templateSource: string
): Promise<ITemplateInfo> {
  try {
    return await Excel.run(async (context: Excel.RequestContext) => {
      // Step 1: Get the template cell
      const templateCell = await getTemplateCell(context, templateSource);

      // Step 2: Validate and extract the formula
      const formula = await validateAndExtractFormula(context, templateCell);

      // Step 3: Return template information
      return {
        cell: templateCell,
        formulaR1C1: formula
      };
    });
  } catch (error) {
    console.error('Error detecting template formula:', error);
    throw error;
  }
}

/**
 * Determines if a cell is blank (no formula AND no value)
 *
 * @param value - Cell value from Excel
 * @param formula - Cell formula from Excel
 * @returns true if cell is truly empty (no formula AND no value)
 */
export function isBlank(value: any, formula: any): boolean {
  // Cell is blank if it has no value AND no formula
  const hasNoValue = value === null || value === undefined || value === '';
  const hasNoFormula = formula === '' || formula === null;

  return hasNoValue && hasNoFormula;
}

/**
 * Determines if a cell contains an Excel error
 *
 * @param value - Cell value from Excel
 * @returns true if cell value is an Excel error type
 */
export function isError(value: any): boolean {
  // Check if value is null or undefined
  if (value === null || value === undefined) {
    return false;
  }

  // Check if value is a string starting with "#"
  // Excel errors: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
  if (typeof value === 'string') {
    // Check if it starts with "#" and matches known error patterns
    const errorPattern = /^#(N\/A|VALUE!|REF!|DIV\/0!|NUM!|NAME\?|NULL!)$/;
    return errorPattern.test(value);
  }

  // Check if value is an Excel error object (Office.js may return errors as objects)
  if (typeof value === 'object' && value.errorType !== undefined) {
    return true;
  }

  return false;
}

/**
 * Determines if a cell is eligible for formula filling based on target condition
 *
 * @param value - Cell value from Excel
 * @param formula - Cell formula from Excel
 * @param targetCondition - Target condition: "blanks", "errors", or "both"
 * @returns true if cell is eligible for formula filling
 */
export function isCellEligible(
  value: any,
  formula: any,
  targetCondition: string
): boolean {
  if (targetCondition === 'blanks') {
    return isBlank(value, formula);
  } else if (targetCondition === 'errors') {
    return isError(value);
  } else if (targetCondition === 'both') {
    return isBlank(value, formula) || isError(value);
  }

  // Unknown target condition, cell is not eligible
  return false;
}

/**
 * Identifies all eligible cells in the selected range
 *
 * @param targetCondition - Target condition: "blanks", "errors", or "both"
 * @returns Promise resolving to array of eligible cell coordinates
 */
export async function identifyEligibleCells(
  targetCondition: string
): Promise<Array<{ row: number; col: number }>> {
  try {
    return await Excel.run(async (context: Excel.RequestContext) => {
      // Get the selected range
      const selectedRange = context.workbook.getSelectedRange();

      // Load range properties in a single batch
      selectedRange.load('values, formulas, rowCount, columnCount');

      // Sync to get the actual values
      await context.sync();

      // Build array of eligible cell coordinates
      const eligibleCells: Array<{ row: number; col: number }> = [];

      // Iterate through all cells in the range
      for (let row = 0; row < selectedRange.rowCount; row++) {
        for (let col = 0; col < selectedRange.columnCount; col++) {
          // Get cell value and formula
          const cellValue = selectedRange.values[row][col];
          const cellFormula = selectedRange.formulas[row][col];

          // Check if cell is eligible
          if (isCellEligible(cellValue, cellFormula, targetCondition)) {
            eligibleCells.push({ row, col });
          }
        }
      }

      return eligibleCells;
    });
  } catch (error) {
    console.error('Error identifying eligible cells:', error);
    throw error;
  }
}

/**
 * Writes formulas to eligible cells in the selected range
 *
 * This function implements batch formula writing with a single context.sync call
 * for optimal performance. R1C1 formulas automatically adjust relative references
 * based on each cell's position.
 *
 * @param eligibleCells - Array of cell coordinates to write formulas to
 * @param templateFormula - R1C1 formula string to write to each cell
 * @returns Promise resolving to the number of cells modified
 */
export async function writeFormulasToEligibleCells(
  eligibleCells: Array<{ row: number; col: number }>,
  templateFormula: string
): Promise<number> {
  // Handle edge case: no eligible cells
  if (eligibleCells.length === 0) {
    return 0;
  }

  try {
    return await Excel.run(async (context: Excel.RequestContext) => {
      // Get the selected range
      const selectedRange = context.workbook.getSelectedRange();

      // Iterate through eligible cells and set formulas
      for (const cell of eligibleCells) {
        // Get the specific cell within the range
        const targetCell = selectedRange.getCell(cell.row, cell.col);

        // Set the formula in R1C1 notation (2D array format)
        targetCell.formulasR1C1 = [[templateFormula]];
      }

      // Sync once after all formulas are set
      await context.sync();

      // Return count of modified cells
      return eligibleCells.length;
    });
  } catch (error) {
    console.error('Error writing formulas to eligible cells:', error);
    throw error;
  }
}

/**
 * Main fill operation orchestrator
 *
 * Executes the complete fill operation pipeline:
 * 1. Validates preflight conditions (template cell has formula)
 * 2. Detects template formula from specified source
 * 3. Identifies eligible cells based on target condition
 * 4. Writes formulas to eligible cells only
 * 5. Returns result with modification count
 *
 * This is the main entry point for the fill operation called from the task pane UI.
 *
 * @param options - Fill operation options (targetCondition, templateSource)
 * @returns Promise resolving to fill operation result
 */
export async function executeFillOperation(
  options: IFillOptions
): Promise<IFillResult> {
  try {
    // Step 1: Validate preflight conditions
    const validationResult = await validatePreflightConditions(options.templateSource);

    if (!validationResult.valid) {
      return {
        modifiedCount: 0,
        success: false,
        error: validationResult.error
      };
    }

    // Step 2: Detect and validate template formula
    const templateInfo = await detectTemplateFormula(options.templateSource);

    // Step 3: Identify eligible cells based on target condition
    const eligibleCells = await identifyEligibleCells(options.targetCondition);

    // Step 4: Check for no eligible cells and return informational message
    if (eligibleCells.length === 0) {
      return {
        modifiedCount: 0,
        success: false,
        error: getNoEligibleCellsMessage(options.targetCondition)
      };
    }

    // Step 5: Write formulas to eligible cells
    const modifiedCount = await writeFormulasToEligibleCells(
      eligibleCells,
      templateInfo.formulaR1C1
    );

    // Step 6: Return success result
    return {
      modifiedCount,
      success: true
    };
  } catch (error) {
    // Handle errors gracefully and return error result
    console.error('Error executing fill operation:', error);

    return {
      modifiedCount: 0,
      success: false,
      error: error instanceof Error ? error.message : 'Unknown error occurred'
    };
  }
}
