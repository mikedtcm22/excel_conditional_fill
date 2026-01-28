# FillGaps v0 Acceptance Checklist

This document provides a comprehensive manual testing checklist for validating the FillGaps Excel add-in before v0 release.

> **Note for Automated Testing:** This checklist requires human execution. The tests cannot be run by automated tools as they require sideloading the add-in in Excel and interacting with the actual Excel UI. A human tester must execute these steps on a Mac with Excel for Microsoft 365.

## Prerequisites

- macOS with Microsoft Excel for Mac (Microsoft 365)
- FillGaps add-in sideloaded successfully (see SIDELOAD.md for instructions)
- Test workbook with sample data prepared

---

## Setup Instructions

### Sideloading on Mac (Microsoft 365)

1. Open Finder and navigate to: `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
2. If the `wef` folder does not exist, create it
3. Copy the `manifest.xml` file from this repository into the `wef` folder
4. Restart Excel completely (Cmd+Q, then reopen)
5. The FillGaps add-in should now appear in the Home ribbon tab

### Creating Test Data

For each test scenario, create fresh test data as specified. After each test, use Cmd+Z to undo changes or close the workbook without saving.

---

## Test Sections

### Section 1: Setup Verification

| ID | Test Case | Steps | Expected Result | Pass/Fail |
|----|-----------|-------|-----------------|-----------|
| 1.1 | Ribbon group appears | Open Excel, check Home ribbon tab | FillGaps group is visible in the Home ribbon | [ ] |
| 1.2 | All buttons visible | Inspect FillGaps ribbon group | Three buttons visible: "Fill Gaps...", "Fill Blanks", "Fill Errors" | [ ] |
| 1.3 | Task pane opens | Click "Fill Gaps..." button | Task pane opens with settings form | [ ] |
| 1.4 | Task pane controls | Inspect task pane | Contains: Template source dropdown, Target condition dropdown, Run button, Status area | [ ] |

---

### Section 2: Task Pane Fill Operation Tests

| ID | Test Case | Setup | Action | Expected Result | Pass/Fail |
|----|-----------|-------|--------|-----------------|-----------|
| 2.1 | Fill blanks via task pane | A1: `=ROW()`, A2-A5: empty | Select A1:A5, open task pane, select "Active cell formula" and "Blanks only", click Run | A2-A5 filled with formula `=ROW()`, status shows "Filled 4 cells successfully" | [ ] |
| 2.2 | Fill errors via task pane | A1: `=VLOOKUP(ROW(),B:C,2,FALSE)`, A2-A3: #N/A errors, A4-A5: values | Select A1:A5, open task pane, select "Active cell formula" and "Errors only", click Run | Only A2-A3 filled, status shows "Filled 2 cells successfully" | [ ] |
| 2.3 | Fill both via task pane | A1: `=ROW()*2`, A2: blank, A3: #VALUE!, A4: value 100 | Select A1:A4, open task pane, select "Top-left cell formula" and "Blanks and errors", click Run | A2 and A3 filled, A4 unchanged, status shows "Filled 2 cells successfully" | [ ] |
| 2.4 | Status area shows count | Any successful fill operation | Complete any fill operation | Status area shows "Filled N cells successfully" with appropriate count | [ ] |
| 2.5 | 2D range fill | A1: `=ROW()+COLUMN()`, A2:C3: mix of blanks | Select A1:C3, open task pane, select "Blanks only", click Run | All blank cells filled with formula | [ ] |

---

### Section 3: Quick Action - Fill Blanks Tests

| ID | Test Case | Setup | Action | Expected Result | Pass/Fail |
|----|-----------|-------|--------|-----------------|-----------|
| 3.1 | Basic fill blanks | A1: `=ROW()`, A2-A5: empty | Select A1:A5 with A1 as active cell, click "Fill Blanks" | A2-A5 filled with formula `=ROW()`, showing values 2,3,4,5 | [ ] |
| 3.2 | Blanks only - values preserved | A1: `=ROW()`, A2: empty, A3: 100, A4: empty, A5: 200 | Select A1:A5 with A1 as active cell, click "Fill Blanks" | A2 and A4 filled, A3 and A5 unchanged | [ ] |
| 3.3 | Blanks only - formulas preserved | A1: `=1`, A2: empty, A3: `=A1+1` | Select A1:A3 with A1 as active cell, click "Fill Blanks" | A2 filled, A3 unchanged (formula preserved) | [ ] |
| 3.4 | No blanks present | A1: `=ROW()`, A2-A5: numeric values | Select A1:A5 with A1 as active cell, click "Fill Blanks" | Alert: "No blank cells found in selection" | [ ] |
| 3.5 | Active cell not A1 | B3: `=ROW()`, B4-B6: empty | Select B3:B6 with B3 as active cell, click "Fill Blanks" | B4-B6 filled with formula | [ ] |

---

### Section 4: Quick Action - Fill Errors Tests

| ID | Test Case | Setup | Action | Expected Result | Pass/Fail |
|----|-----------|-------|--------|-----------------|-----------|
| 4.1 | Basic fill errors | A1: `=VLOOKUP(ROW(),B:C,2,FALSE)`, A2-A3: #N/A | Select A1:A3 with A1 as active cell, click "Fill Errors" | A2-A3 filled with VLOOKUP formula | [ ] |
| 4.2 | Errors only - values preserved | A1: `=ROW()`, A2: #N/A, A3: 100, A4: #VALUE!, A5: 200 | Select A1:A5 with A1 as active cell, click "Fill Errors" | A2 and A4 filled, A3 and A5 unchanged | [ ] |
| 4.3 | Errors only - blanks not filled | A1: `=ROW()`, A2: #N/A, A3: empty, A4: #REF! | Select A1:A4 with A1 as active cell, click "Fill Errors" | A2 and A4 filled, A3 (blank) unchanged | [ ] |
| 4.4 | No errors present | A1: `=ROW()`, A2-A5: numeric values | Select A1:A5 with A1 as active cell, click "Fill Errors" | Alert: "No error cells found in selection" | [ ] |
| 4.5 | Mixed error types | A1: `=ROW()`, A2: #N/A, A3: #VALUE!, A4: #DIV/0! | Select A1:A4 with A1 as active cell, click "Fill Errors" | All error cells (A2-A4) filled | [ ] |

---

### Section 5: Validation - No Formula in Active Cell

| ID | Test Case | Setup | Action | Expected Result | Pass/Fail |
|----|-----------|-------|--------|-----------------|-----------|
| 5.1 | Value in active cell (Fill Blanks) | A1: value 100, A2-A5: empty | Select A1:A5 with A1 as active cell, click "Fill Blanks" | Alert: "Active cell must contain a formula" | [ ] |
| 5.2 | Value in active cell (Fill Errors) | A1: value 100, A2-A3: #N/A errors | Select A1:A3 with A1 as active cell, click "Fill Errors" | Alert: "Active cell must contain a formula" | [ ] |
| 5.3 | Empty active cell (Fill Blanks) | A1: empty, A2-A5: empty | Select A1:A5 with A1 as active cell, click "Fill Blanks" | Alert: "Active cell must contain a formula" | [ ] |
| 5.4 | Task pane validation | A1: value 100, A2-A5: empty | Select A1:A5, open task pane, select "Active cell formula" and "Blanks only", click Run | Status shows error: "Active cell must contain a formula" | [ ] |
| 5.5 | No cells changed after error | A1: value 100, A2: empty (mark value in A2 before test) | Select A1:A2, click "Fill Blanks", dismiss alert | A2 remains unchanged (still empty) | [ ] |

---

### Section 6: Validation - No Eligible Cells

| ID | Test Case | Setup | Action | Expected Result | Pass/Fail |
|----|-----------|-------|--------|-----------------|-----------|
| 6.1 | No blanks (via quick action) | A1: `=ROW()`, A2-A5: numeric values | Select A1:A5 with A1 as active cell, click "Fill Blanks" | Alert: "No blank cells found in selection" | [ ] |
| 6.2 | No errors (via quick action) | A1: `=ROW()`, A2-A5: numeric values | Select A1:A5 with A1 as active cell, click "Fill Errors" | Alert: "No error cells found in selection" | [ ] |
| 6.3 | No blanks (via task pane) | A1: `=ROW()`, A2-A5: numeric values | Open task pane, select "Blanks only", click Run | Status shows: "No blank cells found in selection" | [ ] |
| 6.4 | No errors (via task pane) | A1: `=ROW()`, A2-A5: numeric values | Open task pane, select "Errors only", click Run | Status shows: "No error cells found in selection" | [ ] |

---

### Section 7: Edge Case Scenarios

| ID | Test Case | Setup | Action | Expected Result | Pass/Fail |
|----|-----------|-------|--------|-----------------|-----------|
| 7.1 | Single cell selection | A1: `=ROW()` (formula only, no blanks in selection) | Select only A1, click "Fill Blanks" | Alert: "No blank cells found in selection" (or completes with 0 cells) | [ ] |
| 7.2 | Large range (500+ cells) | A1: `=ROW()`, A2:A500 some blank, some values | Select A1:A500, click "Fill Blanks" | Operation completes within reasonable time (<5 seconds) | [ ] |
| 7.3 | Error type: #N/A | A1: `=ROW()`, A2: `=NA()` | Select A1:A2, click "Fill Errors" | A2 is filled (error detected and replaced) | [ ] |
| 7.4 | Error type: #VALUE! | A1: `=ROW()`, A2: `=1+"text"` | Select A1:A2, click "Fill Errors" | A2 is filled | [ ] |
| 7.5 | Error type: #REF! | A1: `=ROW()`, A2: `=INDIRECT("ZZZ999999")` or create via deleted reference | Select A1:A2, click "Fill Errors" | A2 is filled | [ ] |
| 7.6 | Error type: #DIV/0! | A1: `=ROW()`, A2: `=1/0` | Select A1:A2, click "Fill Errors" | A2 is filled | [ ] |
| 7.7 | Error type: #NUM! | A1: `=ROW()`, A2: `=SQRT(-1)` | Select A1:A2, click "Fill Errors" | A2 is filled | [ ] |
| 7.8 | Error type: #NAME? | A1: `=ROW()`, A2: `=unknownfunction()` | Select A1:A2, click "Fill Errors" | A2 is filled | [ ] |
| 7.9 | Error type: #NULL! | A1: `=ROW()`, A2: `=SUM(A1 B1)` (space instead of colon/comma) | Select A1:A2, click "Fill Errors" | A2 is filled | [ ] |
| 7.10 | Relative reference adjustment | A1: `=A1*2` formula using relative ref, B1: empty, C1: empty | Select A1:C1, click "Fill Blanks" | B1 contains `=B1*2`, C1 contains `=C1*2` (relative refs adjusted) | [ ] |
| 7.11 | Absolute reference preservation | A1: `=$A$1*2`, B1: empty | Select A1:B1, click "Fill Blanks" | B1 contains formula with `$A$1` (absolute ref preserved) | [ ] |

---

## Test Summary

| Section | Total Tests | Passed | Failed |
|---------|-------------|--------|--------|
| 1. Setup Verification | 4 | | |
| 2. Task Pane Fill Operation | 5 | | |
| 3. Quick Action - Fill Blanks | 5 | | |
| 4. Quick Action - Fill Errors | 5 | | |
| 5. Validation - No Formula | 5 | | |
| 6. Validation - No Eligible Cells | 4 | | |
| 7. Edge Cases | 11 | | |
| **TOTAL** | **39** | | |

---

## Critical Path Tests

The following tests are **required to pass** for v0 acceptance:

- [ ] 1.1 - Ribbon group appears
- [ ] 1.2 - All buttons visible
- [ ] 2.1 - Fill blanks via task pane
- [ ] 3.1 - Basic fill blanks (quick action)
- [ ] 4.1 - Basic fill errors (quick action)
- [ ] 5.1 - Validation: No formula in active cell
- [ ] 6.1 - No eligible cells message

---

## Known Issues and Limitations

### v0 Limitations (by design)

1. **Multi-area selections not supported**: Only single contiguous ranges are supported. Selecting multiple non-adjacent ranges (e.g., A1:A5, C1:C5) may produce unexpected results.

2. **Simple alert dialogs**: Error messages are displayed using browser-style `alert()` dialogs rather than styled Office.js dialogs. This is functional but not visually consistent with Office UI.

3. **Silent success for quick actions**: When using "Fill Blanks" or "Fill Errors" ribbon buttons, successful operations complete silently without a notification. Check the cells to verify the operation worked.

4. **No undo integration**: The fill operation does not integrate with Excel's undo stack as a single action. Use Ctrl+Z / Cmd+Z multiple times to undo if needed.

5. **No progress indicator**: For large ranges, there is no visual progress indicator. The operation completes when the button becomes responsive again.

### Features Not Included in v0 (Planned for Future)

- Treat empty string as blank toggle
- Specific error type filtering (choose which error types to fill)
- Convert formulas to values option
- Preview and confirmation dialog before fill
- Settings persistence across sessions
- Undo/redo integration as a single action
- Progress indicator for large ranges

### Workarounds

- **Multi-area selections**: Run the fill operation separately for each contiguous range
- **Verifying quick action success**: Check the cells after clicking the button, or use the task pane for operations where you want status feedback
- **Large ranges**: If the operation appears to hang, wait a few seconds - it may still be processing

---

## Automated Test Coverage

In addition to this manual checklist, the following automated tests exist:

- **Unit tests**: Template detection, cell identification, formula writing
- **Integration tests**: End-to-end fill operation workflows
- **Edge case tests**: Selection edge cases, content edge cases, error type coverage, template formula edge cases
- **Command tests**: Ribbon button command handlers
- **Validation tests**: Preflight validation logic

Run automated tests with: `npm test`

Current test count: **60 tests** (all passing)

---

## Test Execution Log

**Tester:** _____________________

**Date:** _____________________

**Excel Version:** _____________________

**macOS Version:** _____________________

**Notes:**

_____________________________________________________________________________________

_____________________________________________________________________________________

_____________________________________________________________________________________

---

## Sign-off

- [ ] All critical path tests passed
- [ ] No blocking bugs identified
- [ ] v0 ready for release

**Approved by:** _____________________

**Date:** _____________________
