import './taskpane.css';
import { executeFillOperation } from '../fillgaps/engine';
import { IFillOptions } from '../fillgaps/types';

/**
 * FillGaps Task Pane - Event Handlers and Office.js Initialization
 *
 * This module handles:
 * - Office.js initialization and ready state
 * - Form value reading from radio button groups
 * - Run button click handler
 * - Status display updates
 * - Integration with fill operation engine
 */

// Type definitions for form values
type TargetCondition = 'blanks' | 'errors' | 'both';
type TemplateSource = 'activeCell' | 'topLeft';

/**
 * Reads the selected target condition from the radio button group
 * @returns The selected target condition value
 */
function getTargetCondition(): TargetCondition {
  const checkedRadio = document.querySelector<HTMLInputElement>(
    'input[name="targetCondition"]:checked'
  );

  if (!checkedRadio) {
    return 'blanks'; // Default fallback
  }

  return checkedRadio.value as TargetCondition;
}

/**
 * Reads the selected template source from the radio button group
 * @returns The selected template source value
 */
function getTemplateSource(): TemplateSource {
  const checkedRadio = document.querySelector<HTMLInputElement>(
    'input[name="templateSource"]:checked'
  );

  if (!checkedRadio) {
    return 'activeCell'; // Default fallback
  }

  return checkedRadio.value as TemplateSource;
}

/**
 * Updates the status area with a message and optional error styling
 * @param message - The message to display
 * @param isError - Whether to apply error styling (default: false)
 */
function updateStatus(message: string, isError: boolean = false): void {
  const statusArea = document.getElementById('statusArea');

  if (!statusArea) {
    console.error('Status area element not found');
    return;
  }

  // Clear previous status classes
  statusArea.classList.remove('error', 'success');

  // Set message
  statusArea.textContent = message;

  // Apply appropriate styling
  if (isError) {
    statusArea.classList.add('error');
  } else if (message.trim() !== '') {
    statusArea.classList.add('success');
  }
}

/**
 * Executes the fill operation with user-selected options
 *
 * This function integrates with the FillGaps engine to execute the complete
 * fill operation pipeline and display results to the user.
 *
 * @param targetCondition - The target condition selected by user
 * @param templateSource - The template source selected by user
 */
async function runFillOperation(
  targetCondition: TargetCondition,
  templateSource: TemplateSource
): Promise<void> {
  console.log('Fill operation started with:', { targetCondition, templateSource });

  // Clear previous status
  updateStatus('', false);

  // Build options object
  const options: IFillOptions = {
    targetCondition,
    templateSource
  };

  // Execute the fill operation through the engine
  const result = await executeFillOperation(options);

  // Display result in status area
  if (result.success) {
    const message = `Filled ${result.modifiedCount} cell${result.modifiedCount === 1 ? '' : 's'} successfully`;
    updateStatus(message, false);
    console.log('Fill operation completed:', result);
  } else {
    // Display error message
    const errorMessage = result.error || 'Operation failed';
    updateStatus(`Error: ${errorMessage}`, true);
    console.error('Fill operation failed:', result);
  }
}

/**
 * Handles the Run button click event
 * Reads form values and triggers the fill operation
 */
async function handleRunButtonClick(): Promise<void> {
  try {
    // Read form values
    const targetCondition = getTargetCondition();
    const templateSource = getTemplateSource();

    // Log for debugging
    console.log('Run button clicked', { targetCondition, templateSource });

    // Execute fill operation
    await runFillOperation(targetCondition, templateSource);

  } catch (error) {
    console.error('Error during fill operation:', error);
    updateStatus(
      `Error: ${error instanceof Error ? error.message : 'Unknown error occurred'}`,
      true
    );
  }
}

/**
 * Initialize task pane after Office.js is ready
 * Sets up event listeners and prepares UI
 */
function initializeTaskPane(): void {
  console.log('FillGaps add-in loaded successfully');

  // Get Run button
  const runButton = document.getElementById('runButton');

  if (!runButton) {
    console.error('Run button not found in DOM');
    return;
  }

  // Attach click handler
  runButton.addEventListener('click', handleRunButtonClick);

  console.log('Task pane initialized, event listeners attached');
}

/**
 * Office.js initialization
 * Wait for Office to be ready, then initialize task pane
 */
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // DOM should be ready at this point, but use DOMContentLoaded to be safe
    if (document.readyState === 'loading') {
      document.addEventListener('DOMContentLoaded', initializeTaskPane);
    } else {
      // DOM already loaded
      initializeTaskPane();
    }
  }
});
