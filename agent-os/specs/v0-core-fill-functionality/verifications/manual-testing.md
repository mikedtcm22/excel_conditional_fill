# Manual Testing Guide: FillGaps Excel Add-in (v0)

## Overview

This document provides step-by-step instructions for manually testing the FillGaps Excel add-in. Follow these tests in sequence to verify all functionality works as expected.

**Estimated Time:** 30-45 minutes
**Platform:** Excel for Mac (Microsoft 365)
**Prerequisites:** Node.js and npm installed

---

## Part 1: Setup and Installation

### Step 1.1: Install Dependencies

```bash
cd /Users/michaelchristopher/repos/excel_conditional_fill
npm install
```

**Expected Result:** All dependencies install without errors.

---

### Step 1.2: Build the Project

```bash
npm run build
```

**Expected Result:**
- Build completes successfully
- `dist/` folder is created with bundled files
- No TypeScript compilation errors

---

### Step 1.3: Start the Development Server

```bash
npm start
```

**Expected Result:**
- Webpack dev server starts on `https://localhost:3000`
- Console shows "webpack compiled successfully"
- Server continues running (leave this terminal open)

**Note:** If you see SSL certificate warnings, this is expected for localhost development.

---

### Step 1.4: Sideload the Add-in in Excel

In a **new terminal window** (keep the dev server running):

```bash
npm run sideload
```

**Expected Result:**
- Manifest copied to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml`
- Success message displayed

---

### Step 1.5: Load the Add-in in Excel

1. **Open Microsoft Excel** for Mac
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Click **Developer Add-ins** section
4. You should see **FillGaps** in the list
5. Click **FillGaps** to load it
6. Look at the **Home** ribbon tab
7. Verify you see a **FillGaps** group with a **"Fill Gaps..."** button

**Expected Result:** FillGaps ribbon button appears in the Home tab.

**Troubleshooting:**
- If you don't see the add-in, ensure the dev server is running
- Try restarting Excel
- Check that the manifest was copied correctly in Step 1.4

---

## Part 2: Basic Functionality Tests

### Test 2.1: Open Task Pane

**Steps:**
1. Click the **"Fill Gaps..."** button in the FillGaps ribbon group

**Expected Result:**
- Task pane opens on the right side of Excel
- Task pane displays:
  - "FillGaps" header
  - "Target Condition" section with 3 radio buttons
  - "Template Source" section with 2 radio buttons
  - Blue "Run Fill Operation" button
  - Status area below the button

**Verify UI Elements:**
- [X] "Blanks only" radio button is selected by default
- [X] "Active cell formula" radio button is selected by default
- [X] Run button is styled with blue background
- [X] Status area is empty

---

### Test 2.2: Fill Blanks Only - Basic Test

**Setup:**
1. Create a new blank worksheet
2. In cell **A1**, type: `Header`
3. In cell **A2**, type the formula: `=A1&" Value"`
4. In cell **A3**, leave blank (empty cell)
5. In cell **A4**, type: `123`
6. In cell **A5**, leave blank (empty cell)
7. In cell **A6**, type: `Text`

**Test Steps:**
1. Select the range **A2:A6** (click A2, drag to A6)
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area displays: "Filled 2 cells successfully"
- [X] Cell **A3** now contains: `=A2&" Value"` (formula adjusted from A1 to A2)
- [X] Cell **A5** now contains: `=A4&" Value"` (formula adjusted from A1 to A4)
- [X] Cell **A4** (123) is unchanged
- [X] Cell **A6** (Text) is unchanged
- [X] Cell **A2** (original formula) is unchanged

**Verification - R1C1 Reference Adjustment:**
- Click on **A3** and look at the formula bar: Should show `=A2&" Value"`
- Click on **A5** and look at the formula bar: Should show `=A4&" Value"`
- The formula has been correctly adjusted relative to each cell's position

---

### Test 2.3: Fill Errors Only - Basic Test

**Setup:**
1. Create a new worksheet or clear the previous one
2. In cell **A1**, type the formula: `=1/1` (result: 1)
3. In cell **A2**, type the formula: `=1/0` (result: #DIV/0! error)
4. In cell **A3**, type: `42`
5. In cell **A4**, leave blank
6. In cell **A5**, type the formula: `=VLOOKUP(999,A1:A3,2,FALSE)` (result: #N/A error)

**Test Steps:**
1. Click on cell **A1** to make it the active cell
2. Select the range **A2:A5** (you can select A1:A5 if easier)
3. In the FillGaps task pane:
   - Set **Target Condition** to "Errors only"
   - Set **Template Source** to "Active cell formula"
4. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area displays: "Filled 2 cells successfully"
- [X] Cell **A2** (was #DIV/0!) now contains: `=1/1` (copied from A1)
- [X] Cell **A5** (was #N/A) now contains: `=1/1` (copied from A1)
- [X] Cell **A3** (42) is unchanged
- [X] Cell **A4** (blank) is unchanged
- [X] Cell **A1** (original formula) is unchanged

---

### Test 2.4: Fill Blanks + Errors - Combined Mode

**Setup:**
1. Create a new worksheet
2. In cell **B1**, type the formula: `=A1*2`
3. In cell **B2**, leave blank
4. In cell **B3**, type the formula: `=1/0` (result: #DIV/0!)
5. In cell **B4**, leave blank
6. In cell **B5**, type: `100`
7. In cell **B6**, type: `Text`

**Test Steps:**
1. Click on cell **B1** to make it the active cell
2. Select the range **B2:B6**
3. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks + Errors"
   - Set **Template Source** to "Active cell formula"
4. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area displays: "Filled 3 cells successfully"
- [X] Cell **B2** (was blank) now contains a formula
- [X] Cell **B3** (was #DIV/0!) now contains a formula
- [X] Cell **B4** (was blank) now contains a formula
- [X] Cell **B5** (100) is unchanged
- [X] Cell **B6** (Text) is unchanged

---

## Part 3: Template Source Tests

### Test 3.1: Active Cell Formula (Outside Selection)

**Setup:**
1. Create a new worksheet
2. In cell **A1**, type the formula: `=ROW()*10`
3. In cell **B2**, leave blank
4. In cell **B3**, leave blank
5. In cell **B4**, leave blank

**Test Steps:**
1. Click on cell **A1** to make it the active cell
2. Select the range **B2:B4** (without including A1 in selection)
3. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Active cell formula"
4. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area displays: "Filled 3 cells successfully"
- [X] Cell **B2** now contains: `=ROW()*10` (result: 20, since it's row 2)
- [X] Cell **B3** now contains: `=ROW()*10` (result: 30, since it's row 3)
- [X] Cell **B4** now contains: `=ROW()*10` (result: 40, since it's row 4)
- [X] Cell **A1** is unchanged

**Note:** This demonstrates that the active cell can be outside the selection range.

---

### Test 3.2: Top-left Cell in Selection

**Setup:**
1. Create a new worksheet
2. In cell **C1**, type the formula: `=COLUMN()`
3. In cell **C2**, leave blank
4. In cell **C3**, type: `999`
5. In cell **C4**, leave blank

**Test Steps:**
1. Select the range **C1:C4** (include C1 at the top)
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area displays: "Filled 2 cells successfully"
- [X] Cell **C2** now contains: `=COLUMN()` (result: 3, since column C is column 3)
- [X] Cell **C4** now contains: `=COLUMN()` (result: 3)
- [X] Cell **C3** (999) is unchanged
- [X] Cell **C1** (original template) is unchanged

---

## Part 4: Edge Cases and Error Handling

### Test 4.1: No Eligible Cells

**Setup:**
1. Create a new worksheet
2. In cell **D1**, type: `100`
3. In cell **D2**, type: `200`
4. In cell **D3**, type: `300`

**Test Steps:**
1. Select the range **D1:D3**
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area displays: "Filled 0 cells successfully"
- [X] No cells are modified
- [X] All cells still contain their original values (100, 200, 300)

**Note:** This demonstrates graceful handling when no eligible cells are found.

---

### Test 4.2: Template Cell Has No Formula

**Setup:**
1. Create a new worksheet
2. In cell **E1**, type: `Plain Text` (not a formula, just text)
3. In cell **E2**, leave blank
4. In cell **E3**, leave blank

**Test Steps:**
1. Select the range **E1:E3**
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [ ] Status area displays an error message (red text)
- [ ] Error message contains: "Template cell does not contain a formula" or similar
- [ ] No cells are modified
- [ ] Cells **E2** and **E3** remain blank

**Note:** This demonstrates proper error handling for invalid template cells.

**USER NOTE: This is not how this behaved, but what it did I believe is correct.  The blank cells were simply filled in with the "Plain Text" value.  This DOES mirror what would happen if I used the fill feature in Excel on a cell that had a text value in it.**

---

### Test 4.3: All Error Types Detection

**Setup:**
1. Create a new worksheet with lookup data:
   - In **A1:B3**, create a small lookup table:
     - A1: `1`, B1: `Apple`
     - A2: `2`, B2: `Banana`
     - A3: `3`, B3: `Cherry`
2. In cell **C1**, type the formula: `=1/0` (result: #DIV/0!)
3. In cell **C2**, type the formula: `=VLOOKUP(999,A1:B3,2,FALSE)` (result: #N/A)
4. In cell **C3**, type the formula: `=A1/B1` (result: #VALUE! - dividing number by text)
5. In cell **C4**, type the formula: `=SUM(Z999:ZZ9999)` (potentially #REF! if range invalid)
6. Create a template in **D1**: `=1+1` (simple valid formula)

**Test Steps:**
1. Click on cell **D1** to make it the active cell
2. Select the range **C1:C4**
3. In the FillGaps task pane:
   - Set **Target Condition** to "Errors only"
   - Set **Template Source** to "Active cell formula"
4. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area shows cells were filled (at least 3)
- [X] All error cells now contain the formula `=1+1` and display `2`
- [X] The operation successfully detected multiple error types

---

## Part 5: R1C1 Reference Adjustment Tests

### Test 5.1: Relative Row Reference

**Setup:**
1. Create a new worksheet
2. In cell **A1**, type: `10`
3. In cell **A2**, type the formula: `=R[-1]C*2` (references A1, result: 20)
4. In cell **A3**, leave blank
5. In cell **A4**, leave blank
6. In cell **A5**, leave blank

**Test Steps:**
1. Select the range **A2:A5**
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Cell **A3** contains formula that references **A2** (value: 40)
- [X] Cell **A4** contains formula that references **A3** (value: 80)
- [X] Cell **A5** contains formula that references **A4** (value: 160)
- [X] Each cell correctly references the cell immediately above it
- [X] Values double at each step (20 → 40 → 80 → 160)

---

### Test 5.2: Relative Column Reference

**Setup:**
1. Create a new worksheet
2. In cell **A6**, type: `5`
3. In cell **B6**, type the formula: `=RC[-1]+10` (references A6, result: 15)
4. In cell **C6**, leave blank
5. In cell **D6**, leave blank

**Test Steps:**
1. Select the range **B6:D6**
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Cell **C6** contains formula that references **B6** (value: 25)
- [X] Cell **D6** contains formula that references **C6** (value: 35)
- [X] Each cell correctly references the cell immediately to its left
- [X] Values increase by 10 at each step (15 → 25 → 35)

---

### Test 5.3: Absolute Reference (Should NOT Adjust)

**Setup:**
1. Create a new worksheet
2. In cell **A7**, type: `100`
3. In cell **A8**, type the formula: `=R1C1*2` (absolute reference to A1, using R1C1 notation)
4. In cell **A9**, leave blank
5. In cell **A10**, leave blank

**Test Steps:**
1. Select the range **A8:A10**
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Cell **A9** contains formula with absolute reference to **A1** (same as A8)
- [X] Cell **A10** contains formula with absolute reference to **A1** (same as A8)
- [X] All cells reference the same cell (A1), not relative to their position
- [X] If A1 is empty or has value 100, all cells show the same calculated value

**Note:** Absolute references in R1C1 notation (R1C1) remain fixed across all filled cells.

---

## Part 6: Multi-Column Range Test

### Test 6.1: Fill Blanks in a 2D Range

**Setup:**
1. Create a new worksheet
2. Create a small table:
   - **A1**: `Name`, **B1**: `Score`
   - **A2**: `Alice`, **B2**: `=100`
   - **A3**: `Bob`, **B3**: blank
   - **A4**: `Carol`, **B4**: blank
   - **A5**: `Dave`, **B5**: `95`

**Test Steps:**
1. Select the range **B2:B5**
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**

**Expected Results:**
- [X] Status area displays: "Filled 2 cells successfully"
- [X] Cell **B3** now contains: `=100`
- [X] Cell **B4** now contains: `=100`
- [X] Cell **B5** (95) is unchanged
- [X] Cell **B2** (template) is unchanged

---

## Part 7: Performance Test (Optional)

### Test 7.1: Large Range Performance

**Setup:**
1. Create a new worksheet
2. In cell **A1**, type the formula: `=ROW()`
3. Select cells **A2:A100** and delete contents (leave blank)

**Test Steps:**
1. Select the range **A1:A100**
2. In the FillGaps task pane:
   - Set **Target Condition** to "Blanks only"
   - Set **Template Source** to "Top-left cell in selection"
3. Click **"Run Fill Operation"**
4. Note the time it takes to complete

**Expected Results:**
- [X] Operation completes in less than 2 seconds
- [X] Status area displays: "Filled 99 cells successfully"
- [X] Cells A2 through A100 all contain the appropriate formula
- [X] Cell **A2** shows value `2`, cell **A50** shows `50`, cell **A100** shows `100`

**Performance Note:** The spec targets < 2 seconds for ranges up to 1k cells.

---

## Part 8: UI Interaction Tests

### Test 8.1: Radio Button Behavior

**Test Steps:**
1. In the FillGaps task pane, click each radio button option
2. Verify only one option can be selected at a time within each group

**Expected Results:**
- [X] Target Condition radio buttons are mutually exclusive
- [X] Template Source radio buttons are mutually exclusive
- [X] Clicking a radio button deselects the previously selected option in that group
- [X] Radio buttons are clearly visible and easy to click

---

### Test 8.2: Status Area Feedback

**Test Steps:**
1. Run a successful operation (use any test from Part 2)
2. Observe the status area message
3. Run an operation that produces an error (use Test 4.2)
4. Observe the error message

**Expected Results:**
- [X] Success messages appear in normal text (not red)
- [X] Success messages include the count: "Filled X cells successfully"
- [X] Error messages appear in red text
- [X] Error messages clearly describe the problem
- [X] Status area clears when starting a new operation

---

### Test 8.3: Task Pane Responsiveness

**Test Steps:**
1. Click the "Fill Gaps..." button multiple times to close and reopen the task pane
2. Resize the Excel window
3. Scroll the worksheet while the task pane is open

**Expected Results:**
- [X] Task pane opens and closes smoothly
- [X] Task pane maintains its size when reopened
- [X] UI elements in task pane remain visible and properly styled
- [X] Task pane doesn't interfere with Excel's normal operations
- [X] Task pane width is appropriate (300-400px)

---

## Part 9: Cleanup and Shutdown

### Step 9.1: Stop the Development Server

1. Go to the terminal where you ran `npm start`
2. Press **Ctrl+C** to stop the server

**Expected Result:** Server shuts down cleanly.

---

### Step 9.2: Remove the Add-in from Excel (Optional)

If you want to remove the add-in from Excel:

1. Close all Excel windows
2. Delete the manifest file:
   ```bash
   rm ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml
   ```
3. Reopen Excel

**Expected Result:** FillGaps add-in no longer appears in Excel.

---

## Test Results Summary

After completing all tests, fill out this checklist:

### Core Functionality
- [X] Task pane opens and displays correctly
- [X] Blanks only mode works correctly
- [X] Errors only mode works correctly
- [X] Blanks + Errors mode works correctly
- [X] Active cell template source works
- [X] Top-left template source works

### Edge Cases
- [X] Handles no eligible cells gracefully
- [X] Handles missing formula in template cell
- [X] Detects all Excel error types

### R1C1 Reference Adjustment
- [X] Relative row references adjust correctly
- [X] Relative column references adjust correctly
- [X] Absolute references remain fixed

### User Experience
- [X] Status messages are clear and accurate
- [X] Error messages are helpful
- [X] UI is responsive and intuitive
- [X] Operation completes quickly (< 2 seconds for typical ranges)

### Quality
- [X] No cells with existing values are overwritten
- [X] No cells with existing formulas are overwritten
- [X] No console errors appear during normal operation

---

## Troubleshooting

### Issue: Add-in doesn't appear in Excel

**Solutions:**
1. Ensure the dev server is running (`npm start`)
2. Verify the manifest was copied: `ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
3. Run `npm run sideload` again
4. Restart Excel completely
5. Check if Excel trusts the localhost certificate

---

### Issue: Task pane shows blank or error

**Solutions:**
1. Check the dev server terminal for errors
2. Open browser console (if available) to see JavaScript errors
3. Verify the dev server is serving on `https://localhost:3000`
4. Try rebuilding: `npm run build && npm start`

---

### Issue: "Template cell does not contain a formula" error

**Solutions:**
1. Verify the template cell actually contains a formula (starts with `=`)
2. Check that the correct Template Source option is selected
3. If using "Active cell formula", ensure the active cell has a formula
4. If using "Top-left cell in selection", ensure the first cell in selection has a formula

---

### Issue: Formulas aren't adjusting correctly

**Solutions:**
1. This is expected behavior - Excel automatically adjusts R1C1 formulas
2. Check if you're using absolute references (R1C1) instead of relative (R[-1]C)
3. Verify you're looking at the formula bar to see the actual formula, not just the result

---

## Success Criteria

The implementation passes manual testing if:

1. ✅ All tests in Parts 2-6 pass
2. ✅ No unexpected errors occur
3. ✅ Status messages are accurate
4. ✅ R1C1 references adjust correctly
5. ✅ Edge cases are handled gracefully
6. ✅ No existing cells are overwritten
7. ✅ Operations complete in < 2 seconds for typical ranges

---

## Notes for Future Testing

- This test suite covers v0 functionality only
- Future versions may include additional features (see roadmap)
- Report any bugs or unexpected behavior to the development team
- Document any new edge cases discovered during testing

---

**Testing Date:** _______________
**Tester Name:** _______________
**Excel Version:** _______________
**macOS Version:** _______________
**Test Result:** ☐ Pass  ☐ Fail (with notes)
