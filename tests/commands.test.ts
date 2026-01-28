/**
 * Quick Action Command Tests
 *
 * Focused tests for ribbon command handlers (Task Group 9)
 * Tests command execution with correct options and event.completed() calls
 */

// Mock executeFillOperation before importing commands
const mockExecuteFillOperation = jest.fn();
const mockShowErrorDialog = jest.fn();

jest.mock('../src/fillgaps/engine', () => ({
  executeFillOperation: mockExecuteFillOperation
}));

jest.mock('../src/fillgaps/validation', () => ({
  showErrorDialog: mockShowErrorDialog
}));

// Mock Office global
const mockCompleted = jest.fn();
const mockEvent: Office.AddinCommands.Event = {
  completed: mockCompleted,
  source: { id: 'testButton' }
};

// Setup Office.actions mock
const mockAssociate = jest.fn();
(global as any).Office = {
  onReady: jest.fn((callback: any) => callback()),
  actions: {
    associate: mockAssociate
  }
};

describe('Quick Action Commands', () => {
  beforeEach(() => {
    jest.clearAllMocks();
    // Reset module cache to ensure fresh imports
    jest.resetModules();
  });

  /**
   * Test 1: fillBlanksCommand executes with correct options
   */
  test('fillBlanksCommand executes with templateSource: activeCell, targetCondition: blanks', async () => {
    // Setup mock to resolve successfully
    mockExecuteFillOperation.mockResolvedValue({
      success: true,
      modifiedCount: 5
    });

    // Import and get the command handler
    const { fillBlanksCommand } = await import('../src/commands/commands');

    // Execute the command
    await fillBlanksCommand(mockEvent);

    // Verify executeFillOperation was called with correct options
    expect(mockExecuteFillOperation).toHaveBeenCalledTimes(1);
    expect(mockExecuteFillOperation).toHaveBeenCalledWith({
      templateSource: 'activeCell',
      targetCondition: 'blanks'
    });
  });

  /**
   * Test 2: fillErrorsCommand executes with correct options
   */
  test('fillErrorsCommand executes with templateSource: activeCell, targetCondition: errors', async () => {
    // Setup mock to resolve successfully
    mockExecuteFillOperation.mockResolvedValue({
      success: true,
      modifiedCount: 3
    });

    // Import and get the command handler
    const { fillErrorsCommand } = await import('../src/commands/commands');

    // Execute the command
    await fillErrorsCommand(mockEvent);

    // Verify executeFillOperation was called with correct options
    expect(mockExecuteFillOperation).toHaveBeenCalledTimes(1);
    expect(mockExecuteFillOperation).toHaveBeenCalledWith({
      templateSource: 'activeCell',
      targetCondition: 'errors'
    });
  });

  /**
   * Test 3: event.completed() is called after successful command execution
   */
  test('event.completed() is called after fillBlanksCommand execution', async () => {
    // Setup mock to resolve successfully
    mockExecuteFillOperation.mockResolvedValue({
      success: true,
      modifiedCount: 2
    });

    // Import and get the command handler
    const { fillBlanksCommand } = await import('../src/commands/commands');

    // Execute the command
    await fillBlanksCommand(mockEvent);

    // Verify event.completed() was called
    expect(mockCompleted).toHaveBeenCalledTimes(1);
  });

  /**
   * Test 4: event.completed() is called even after command error and showErrorDialog is used
   */
  test('event.completed() is called after fillErrorsCommand error and showErrorDialog displays message', async () => {
    // Setup mock to resolve with error
    mockExecuteFillOperation.mockResolvedValue({
      success: false,
      modifiedCount: 0,
      error: 'Active cell must contain a formula'
    });

    // Import and get the command handler
    const { fillErrorsCommand } = await import('../src/commands/commands');

    // Execute the command
    await fillErrorsCommand(mockEvent);

    // Verify event.completed() was called even after error
    expect(mockCompleted).toHaveBeenCalledTimes(1);

    // Verify showErrorDialog was called with error message
    expect(mockShowErrorDialog).toHaveBeenCalledWith('Active cell must contain a formula');
  });
});
