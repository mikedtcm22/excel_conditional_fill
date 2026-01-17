/// <reference types="office-js" />

/**
 * FillGaps Type Definitions
 *
 * TypeScript interfaces and types for the FillGaps engine module
 */

/**
 * Options for the fill operation
 */
export interface IFillOptions {
  /**
   * Target condition for identifying eligible cells
   * - "blanks": Only blank cells (no formula AND no value)
   * - "errors": Only cells with Excel errors (#N/A, #VALUE!, etc.)
   * - "both": Both blanks and errors
   */
  targetCondition: string;

  /**
   * Source of the template formula
   * - "activeCell": Use the currently active cell's formula
   * - "topLeft": Use the top-left cell of the selection
   */
  templateSource: string;
}

/**
 * Information about the detected template cell and its formula
 */
export interface ITemplateInfo {
  /**
   * The Excel range object representing the template cell
   */
  cell: Excel.Range;

  /**
   * The R1C1 formula string extracted from the template cell
   */
  formulaR1C1: string;
}

/**
 * Result of a fill operation
 */
export interface IFillResult {
  /**
   * Number of cells that were modified during the operation
   */
  modifiedCount: number;

  /**
   * Whether the operation completed successfully
   */
  success: boolean;

  /**
   * Error message if the operation failed (optional)
   */
  error?: string;
}
