/**
 * Cell Identification Tests
 *
 * Focused tests for eligible cell identification module (Task Group 6)
 * Tests critical eligibility detection behaviors only
 */

describe('Eligible Cell Identification', () => {
  /**
   * Test 1: Blank detection - truly empty cell
   */
  test('isBlank returns true for cell with no value and no formula', async () => {
    const { isBlank } = await import('../src/fillgaps/engine');

    // Truly empty cell: no value AND no formula
    expect(isBlank(null, '')).toBe(true);
    expect(isBlank(undefined, '')).toBe(true);
    expect(isBlank('', '')).toBe(true);
    expect(isBlank(null, null)).toBe(true);
    expect(isBlank(undefined, null)).toBe(true);
    expect(isBlank('', null)).toBe(true);
  });

  /**
   * Test 2: Blank detection - cell with value is not blank
   */
  test('isBlank returns false for cell with value', async () => {
    const { isBlank } = await import('../src/fillgaps/engine');

    // Cell has value, not blank
    expect(isBlank(100, '')).toBe(false);
    expect(isBlank('text', '')).toBe(false);
    expect(isBlank(0, '')).toBe(false);
  });

  /**
   * Test 3: Blank detection - cell with formula is not blank
   */
  test('isBlank returns false for cell with formula', async () => {
    const { isBlank } = await import('../src/fillgaps/engine');

    // Cell has formula, not blank (even if formula returns empty)
    expect(isBlank('', '=A1')).toBe(false);
    expect(isBlank(null, '=SUM(A1:A5)')).toBe(false);
    expect(isBlank(0, '=R[-1]C')).toBe(false);
  });

  /**
   * Test 4: Error detection - Excel error types
   */
  test('isError returns true for Excel error values', async () => {
    const { isError } = await import('../src/fillgaps/engine');

    // String-based error detection (Excel returns errors as strings)
    expect(isError('#N/A')).toBe(true);
    expect(isError('#VALUE!')).toBe(true);
    expect(isError('#REF!')).toBe(true);
    expect(isError('#DIV/0!')).toBe(true);
    expect(isError('#NUM!')).toBe(true);
    expect(isError('#NAME?')).toBe(true);
    expect(isError('#NULL!')).toBe(true);
  });

  /**
   * Test 5: Error detection - non-error values
   */
  test('isError returns false for non-error values', async () => {
    const { isError } = await import('../src/fillgaps/engine');

    // Regular values are not errors
    expect(isError(100)).toBe(false);
    expect(isError('text')).toBe(false);
    expect(isError('')).toBe(false);
    expect(isError(null)).toBe(false);
    expect(isError(undefined)).toBe(false);
    expect(isError('#hashtag')).toBe(false); // Not a valid Excel error
  });

  /**
   * Test 6: Cell eligibility - blanks mode
   */
  test('isCellEligible returns true for blank cells when targetCondition is blanks', async () => {
    const { isCellEligible } = await import('../src/fillgaps/engine');

    // Blanks mode: only blank cells are eligible
    expect(isCellEligible(null, '', 'blanks')).toBe(true);
    expect(isCellEligible('', null, 'blanks')).toBe(true);
    expect(isCellEligible(100, '', 'blanks')).toBe(false);
    expect(isCellEligible('#N/A', '', 'blanks')).toBe(false);
  });

  /**
   * Test 7: Cell eligibility - errors mode
   */
  test('isCellEligible returns true for error cells when targetCondition is errors', async () => {
    const { isCellEligible } = await import('../src/fillgaps/engine');

    // Errors mode: only error cells are eligible
    expect(isCellEligible('#N/A', '', 'errors')).toBe(true);
    expect(isCellEligible('#VALUE!', '=VLOOKUP()', 'errors')).toBe(true);
    expect(isCellEligible(null, '', 'errors')).toBe(false);
    expect(isCellEligible(100, '', 'errors')).toBe(false);
  });

  /**
   * Test 8: Cell eligibility - both mode
   */
  test('isCellEligible returns true for both blanks and errors when targetCondition is both', async () => {
    const { isCellEligible } = await import('../src/fillgaps/engine');

    // Both mode: blanks OR errors are eligible
    expect(isCellEligible(null, '', 'both')).toBe(true);
    expect(isCellEligible('#N/A', '', 'both')).toBe(true);
    expect(isCellEligible('#REF!', '=A1', 'both')).toBe(true);
    expect(isCellEligible(100, '', 'both')).toBe(false);
    expect(isCellEligible('text', '=A1', 'both')).toBe(false);
  });
});
