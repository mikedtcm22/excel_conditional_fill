/// <reference types="office-js" />

/**
 * FillGaps Validation Module
 *
 * Preflight validation and error display functions for the FillGaps add-in.
 * Validates that template cells contain formulas and displays appropriate error messages.
 */

/**
 * Result of a preflight validation check
 */
export interface IValidationResult {
  /**
   * Whether the validation passed
   */
  valid: boolean;

  /**
   * Error message if validation failed (optional)
   */
  error?: string;
}

/**
 * Validates preflight conditions before executing a fill operation
 *
 * Checks that the template cell (based on templateSource) contains a formula.
 * Returns validation result with error message if validation fails.
 *
 * @param templateSource - Source of template: "activeCell" or "topLeft"
 * @returns Promise resolving to validation result
 */
export async function validatePreflightConditions(
  templateSource: string
): Promise<IValidationResult> {
  try {
    return await Excel.run(async (context: Excel.RequestContext) => {
      let templateCell: Excel.Range;

      // Get template cell based on templateSource
      if (templateSource === 'activeCell') {
        templateCell = context.workbook.getActiveCell();
      } else if (templateSource === 'topLeft') {
        const selectedRange = context.workbook.getSelectedRange();
        templateCell = selectedRange.getCell(0, 0);
      } else {
        return {
          valid: false,
          error: `Invalid template source: ${templateSource}`
        };
      }

      // Load formulasR1C1 property
      templateCell.load('formulasR1C1');

      // Sync to get actual values
      await context.sync();

      // Check if cell has a formula (not empty/null)
      const formula = templateCell.formulasR1C1[0][0];

      if (!formula || formula === '') {
        return {
          valid: false,
          error: 'Active cell must contain a formula'
        };
      }

      // Validation passed
      return {
        valid: true
      };
    });
  } catch (error) {
    console.error('Error during preflight validation:', error);
    return {
      valid: false,
      error: error instanceof Error ? error.message : 'Validation error occurred'
    };
  }
}

/**
 * Displays an error message in a modal dialog
 *
 * Uses simple alert() for v0 (cross-platform compatible).
 * Future enhancement: Office.context.ui.displayDialogAsync for styled dialogs.
 *
 * @param message - The error message to display
 */
export function showErrorDialog(message: string): void {
  alert(message);
}

/**
 * Displays an informational message
 *
 * For task pane context, this function would update the status area.
 * Note: This function is provided for completeness but task pane uses
 * its own updateStatus function directly.
 *
 * @param message - The informational message to display
 */
export function showInfoMessage(message: string): void {
  // For quick actions, we use silent completion (no notification)
  // This function is a placeholder for future enhancements
  console.log('Info:', message);
}

/**
 * Gets the appropriate "no eligible cells" error message based on target condition
 *
 * @param targetCondition - Target condition: "blanks", "errors", or "both"
 * @returns The appropriate error message string
 */
export function getNoEligibleCellsMessage(targetCondition: string): string {
  switch (targetCondition) {
    case 'blanks':
      return 'No blank cells found in selection';
    case 'errors':
      return 'No error cells found in selection';
    case 'both':
      return 'No blank or error cells found in selection';
    default:
      return 'No eligible cells found in selection';
  }
}
