/// <reference types="office-js" />

/**
 * FillGaps Ribbon Command Handlers
 *
 * Quick action commands for one-click fill operations from the ribbon.
 * These commands execute fill operations directly without opening the task pane.
 */

import { executeFillOperation } from '../fillgaps/engine';
import { showErrorDialog } from '../fillgaps/validation';

/**
 * Fill Blanks Command Handler
 *
 * Executes a fill operation targeting only blank cells using the active cell's formula.
 * Called when the user clicks the "Fill Blanks" ribbon button.
 *
 * @param event - Office add-in command event
 */
export async function fillBlanksCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const result = await executeFillOperation({
      templateSource: 'activeCell',
      targetCondition: 'blanks'
    });

    if (!result.success && result.error) {
      showErrorDialog(result.error);
    }
    // Success case: silent completion (no notification for v0)
  } catch (error) {
    showErrorDialog('An unexpected error occurred');
    console.error('Fill Blanks command error:', error);
  } finally {
    // Always call event.completed() to signal Office the command finished
    event.completed();
  }
}

/**
 * Fill Errors Command Handler
 *
 * Executes a fill operation targeting only error cells using the active cell's formula.
 * Called when the user clicks the "Fill Errors" ribbon button.
 *
 * @param event - Office add-in command event
 */
export async function fillErrorsCommand(event: Office.AddinCommands.Event): Promise<void> {
  try {
    const result = await executeFillOperation({
      templateSource: 'activeCell',
      targetCondition: 'errors'
    });

    if (!result.success && result.error) {
      showErrorDialog(result.error);
    }
    // Success case: silent completion (no notification for v0)
  } catch (error) {
    showErrorDialog('An unexpected error occurred');
    console.error('Fill Errors command error:', error);
  } finally {
    // Always call event.completed() to signal Office the command finished
    event.completed();
  }
}

/**
 * Office.js initialization and command registration
 */
Office.onReady(() => {
  console.log('FillGaps commands loaded');

  // Register command handlers with Office
  Office.actions.associate('fillBlanksCommand', fillBlanksCommand);
  Office.actions.associate('fillErrorsCommand', fillErrorsCommand);
});
