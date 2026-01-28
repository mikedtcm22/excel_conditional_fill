# Specification: v0 Completion (Quick Actions, Validation, Testing)

## Goal
Complete the remaining v0 roadmap items (6, 7, 8) to deliver a fully functional FillGaps add-in with one-click ribbon actions, proper user-facing validation and error handling, and comprehensive test coverage for acceptance and edge cases.

## User Stories
- As a power user, I want to click "Fill Blanks" in the ribbon and have it immediately fill blanks using my active cell's formula so that I can work faster without opening the task pane
- As a power user, I want to click "Fill Errors" in the ribbon and have it immediately fill error cells using my active cell's formula so that I can quickly fix lookup failures
- As a user who made a mistake, I want to see a clear error dialog when my active cell has no formula so that I understand why the operation failed
- As a user filling gaps, I want to see informational status in the task pane so that I know the operation succeeded and how many cells were modified
- As a developer, I want automated tests covering edge cases so that I can confidently maintain and extend the codebase

## Core Requirements

### Quick Action Commands (Roadmap Item 6)
- Add "Fill Blanks" button to ribbon that executes fill operation with:
  - Template source: Active cell
  - Target condition: Blanks only
  - One-click execution (no task pane required)
- Add "Fill Errors" button to ribbon that executes fill operation with:
  - Template source: Active cell
  - Target condition: Errors only
  - One-click execution (no task pane required)
- Both buttons use existing `executeFillOperation()` engine function

### Validation & Error Messages (Roadmap Item 7)
- Preflight validation checks before execution:
  - Active cell must contain a formula
  - Selection must be a contiguous range (single area)
  - At least one cell in range (non-empty selection)
- Two-tier message display system:
  - Informational messages (success, count) display in task pane status area
  - Error messages display in modal dialog (Office.js dialog or alert)
- Clear, user-facing error messages:
  - "Active cell must contain a formula"
  - "No blank cells found in selection" / "No error cells found in selection"
  - "Selection must be a contiguous range"

### End-to-End Testing (Roadmap Item 8)
- Manual acceptance testing checklist document
- Automated edge case tests covering:
  - Empty selection handling
  - Single cell selection
  - Large range performance (1k+ cells)
  - Mixed content ranges
  - All Excel error types
  - Template formula edge cases

## Visual Design
No visual mockups provided. Implementation follows existing UI patterns:

**Ribbon Updates:**
- Add two new buttons to existing FillGaps ribbon group:
  - "Fill Blanks" button (ExecuteFunction action type)
  - "Fill Errors" button (ExecuteFunction action type)
- Use appropriate icons from Office icon set

**Error Dialog:**
- Use `Office.context.ui.displayDialogAsync` or simple `alert()` for v0
- Display single error message with OK button to dismiss

**Task Pane Status Area:**
- Already exists (`#statusArea` element)
- Continue using existing `updateStatus()` function for informational messages

## Reusable Components

### Existing Code to Leverage

**Engine Module (`src/fillgaps/engine.ts`):**
- `executeFillOperation(options)` - Main fill orchestrator, reuse directly
- `detectTemplateFormula(templateSource)` - Template detection with validation
- `identifyEligibleCells(targetCondition)` - Cell eligibility identification
- `writeFormulasToEligibleCells(cells, formula)` - Batch formula writing
- `isBlank()`, `isError()`, `isCellEligible()` - Cell checking utilities

**Types (`src/fillgaps/types.ts`):**
- `IFillOptions` - Options interface (templateSource, targetCondition)
- `IFillResult` - Result interface (success, modifiedCount, error)
- `ITemplateInfo` - Template info interface

**Task Pane (`src/taskpane/taskpane.ts`):**
- `updateStatus(message, isError)` - Status display function to reuse
- `runFillOperation()` - Fill execution flow to model after

**Manifest (`manifest.xml`):**
- Existing FillGapsGroup ribbon group to extend
- Existing Commands.Url function file reference

**Existing Tests (`tests/`):**
- `integration.test.ts` - Mock patterns for Excel.js API
- Test structure and assertion patterns to follow

### New Components Required

**Quick Action Command Handlers (`src/commands/commands.ts`):**
- `fillBlanksCommand()` - Handler for Fill Blanks ribbon button
- `fillErrorsCommand()` - Handler for Fill Errors ribbon button
- WHY: Current commands.ts only has Office.onReady; needs actual command functions

**Validation Module (`src/fillgaps/validation.ts`):**
- `validatePreflightConditions()` - Check selection and template before execution
- `showErrorDialog(message)` - Display error in modal dialog
- WHY: Validation logic should be separate from engine; error dialogs need new implementation

**Manifest Button Definitions:**
- Add Fill Blanks and Fill Errors button controls to manifest
- WHY: Current manifest only has task pane button; needs ExecuteFunction buttons

**Edge Case Tests (`tests/edge-cases.test.ts`):**
- New test file for comprehensive edge case coverage
- WHY: Existing integration tests cover happy paths; edge cases need dedicated tests

**Manual Test Checklist (`docs/acceptance-checklist.md`):**
- Documented manual testing procedures
- WHY: Manual verification needed for v0 sign-off

## Technical Approach

### Quick Action Commands

**Manifest Changes:**
- Add two new `<Control xsi:type="Button">` elements to FillGapsGroup
- Use `<Action xsi:type="ExecuteFunction">` (not ShowTaskpane)
- Reference function names: `fillBlanksCommand`, `fillErrorsCommand`
- Add corresponding `<FunctionName>` elements to Resources

**Command Handler Pattern:**
```
fillBlanksCommand(event: Office.AddinCommands.Event):
  1. Call executeFillOperation({ targetCondition: 'blanks', templateSource: 'activeCell' })
  2. If error, show modal dialog with error message
  3. If success, show notification or silent completion
  4. Call event.completed() to signal Office that command finished
```

**Global Function Registration:**
- Commands must be registered globally: `Office.actions.associate("fillBlanksCommand", fillBlanksCommand)`
- Function file (commands.html) must load commands.ts bundle

### Validation & Error Messages

**Preflight Validation Flow:**
```
validatePreflightConditions():
  1. Get active cell and check if it has a formula
  2. Get selected range and verify it's a single area (not multi-select)
  3. Return validation result with specific error message if failed
```

**Error Dialog Implementation:**
- Use `alert()` for v0 simplicity (cross-platform compatible)
- Future enhancement: Office.context.ui.displayDialogAsync for styled dialogs

**Integration with Engine:**
- Add validation call at start of `executeFillOperation()` or in command handlers
- Return validation errors in `IFillResult.error` field
- Task pane continues to show success messages in status area
- Quick actions show errors in modal dialog

### End-to-End Testing

**Automated Edge Case Tests:**
- Mock Excel.js API following existing patterns in integration.test.ts
- Test file: `tests/edge-cases.test.ts`
- Test categories:
  - Selection edge cases (empty, single cell, very large)
  - Content edge cases (all blanks, all errors, no eligible cells)
  - Error type coverage (all 7 Excel error types)
  - Template edge cases (complex formulas, absolute refs, mixed refs)

**Manual Acceptance Checklist:**
- Document file: `docs/acceptance-checklist.md`
- Covers all v0 functionality with step-by-step verification
- Includes setup instructions for sideloading

## Out of Scope

**Features NOT in this spec (v0.2 or later):**
- Treat empty string as blank toggle
- Specific error type filtering (checklist)
- Convert to values option
- Preview & confirmation dialog
- Settings persistence across sessions
- Undo/redo integration
- Progress indicator for large ranges

**Technical Constraints:**
- Multi-area selections not supported (single contiguous range only)
- No styled Office.js dialog (using simple alert for v0)
- No notification toast after quick actions (silent success for v0)

## Success Criteria

### Quick Action Commands
- [ ] "Fill Blanks" button appears in FillGaps ribbon group
- [ ] "Fill Errors" button appears in FillGaps ribbon group
- [ ] Clicking "Fill Blanks" fills only blank cells using active cell formula
- [ ] Clicking "Fill Errors" fills only error cells using active cell formula
- [ ] Both commands complete in < 2 seconds for ranges up to 1k cells
- [ ] Error dialog appears if active cell has no formula

### Validation & Error Messages
- [ ] Error dialog displays "Active cell must contain a formula" when appropriate
- [ ] Error dialog displays "No blank cells found in selection" when appropriate
- [ ] Error dialog displays "No error cells found in selection" when appropriate
- [ ] Task pane status shows "Filled N cells successfully" on success
- [ ] Task pane status shows error styling for validation failures

### End-to-End Testing
- [ ] All automated edge case tests pass
- [ ] Manual acceptance checklist document exists
- [ ] Manual testing completed against checklist (all items pass)
- [ ] Test coverage includes all 7 Excel error types
- [ ] Test coverage includes single cell, empty range, and large range scenarios

### Overall v0 Completion
- [ ] All 8 v0 roadmap items marked complete
- [ ] Add-in sideloads successfully on Mac
- [ ] Core fill functionality works end-to-end
- [ ] No known critical bugs

## Acceptance Test Checklist (Manual)

### Setup
1. Sideload add-in on Excel for Mac (Microsoft 365)
2. Verify FillGaps ribbon group appears on Home tab
3. Verify all three buttons visible: "Fill Gaps...", "Fill Blanks", "Fill Errors"

### Task Pane Fill Operation
1. Create test range with formula in A1, blanks in A2:A5
2. Select A1:A5, click "Fill Gaps..." to open task pane
3. Select "Blanks only" and "Active cell formula"
4. Click Run - verify blanks filled, status shows count
5. Undo and repeat with "Errors only" - verify no changes (no errors)
6. Add #N/A error to A3, repeat - verify only A3 filled

### Quick Action: Fill Blanks
1. Create range with formula in A1, blanks in A2:A5
2. Select A1:A5 with A1 as active cell
3. Click "Fill Blanks" ribbon button
4. Verify all blanks filled with A1's formula
5. Verify existing values/formulas unchanged

### Quick Action: Fill Errors
1. Create range with formula in A1, #N/A errors in A2:A3, values in A4:A5
2. Select A1:A5 with A1 as active cell
3. Click "Fill Errors" ribbon button
4. Verify only A2:A3 (errors) filled with A1's formula
5. Verify values in A4:A5 unchanged

### Validation: No Formula in Active Cell
1. Select range with value (not formula) in active cell
2. Click "Fill Blanks" ribbon button
3. Verify error dialog: "Active cell must contain a formula"
4. Dismiss dialog, verify no cells changed

### Validation: No Eligible Cells
1. Create range with formula in A1, values in A2:A5 (no blanks)
2. Select A1:A5, click "Fill Blanks"
3. Verify status message: "No blank cells found in selection"

### Edge Cases
1. Single cell selection with formula - verify operation completes (0 cells modified)
2. Large range (500+ cells) - verify operation completes in reasonable time
3. Various error types (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!) - verify all detected

## Automated Edge Case Test Specifications

### Selection Edge Cases
```
Test: Empty selection (user has no range selected)
Expected: Operation returns error or 0 cells modified

Test: Single cell selection with formula
Expected: Success with 0 cells modified (no eligible cells)

Test: Single cell selection that is blank
Expected: Success with 0 cells modified (blank is also template source)

Test: Large range (1000 cells, 50% blanks)
Expected: Success with ~500 cells modified, completes in < 5 seconds
```

### Content Edge Cases
```
Test: Range where all cells are eligible (all blanks)
Expected: All cells filled except template cell

Test: Range where no cells are eligible (all have values/formulas)
Expected: Success with 0 cells modified

Test: Range with only the template cell containing formula
Expected: Success, non-template blanks/errors filled
```

### Error Type Coverage
```
Test: Cell with #N/A error
Expected: Detected as error, eligible for filling

Test: Cell with #VALUE! error
Expected: Detected as error, eligible for filling

Test: Cell with #REF! error
Expected: Detected as error, eligible for filling

Test: Cell with #DIV/0! error
Expected: Detected as error, eligible for filling

Test: Cell with #NUM! error
Expected: Detected as error, eligible for filling

Test: Cell with #NAME? error
Expected: Detected as error, eligible for filling

Test: Cell with #NULL! error
Expected: Detected as error, eligible for filling
```

### Template Formula Edge Cases
```
Test: Template with simple relative reference (=A1)
Expected: R1C1 formula adjusts correctly per cell position

Test: Template with absolute reference (=$A$1)
Expected: R1C1 formula preserves absolute reference

Test: Template with mixed reference (=$A1 or =A$1)
Expected: R1C1 formula preserves mixed reference behavior

Test: Template with complex formula (=VLOOKUP(A1,Sheet2!A:B,2,FALSE))
Expected: Formula copied correctly to eligible cells

Test: Template with named range reference
Expected: Named range preserved in filled cells
```

## File Changes Summary

### Files to Modify
- `manifest.xml` - Add Fill Blanks and Fill Errors button definitions
- `src/commands/commands.ts` - Add command handler functions
- `src/commands/commands.html` - Ensure commands.ts is bundled and loaded
- `src/fillgaps/engine.ts` - Add validation calls (optional, could be in validation.ts)
- `src/taskpane/taskpane.ts` - Update to use new validation module

### Files to Create
- `src/fillgaps/validation.ts` - Preflight validation and error dialog functions
- `tests/edge-cases.test.ts` - Automated edge case test suite
- `docs/acceptance-checklist.md` - Manual testing checklist (copy from this spec)

### Files Unchanged
- `src/fillgaps/types.ts` - Existing types sufficient
- `src/taskpane/taskpane.html` - No UI changes needed
- `src/taskpane/taskpane.css` - No styling changes needed
- `tests/integration.test.ts` - Existing tests remain valid
