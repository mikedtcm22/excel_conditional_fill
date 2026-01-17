# Task Breakdown: FillGaps Core Fill Functionality (v0)

## Overview
**Total Tasks:** 36 sub-tasks across 8 task groups
**Spec Scope:** Items 1-5 from roadmap (Manifest, Settings Task Pane, Template Detection, Cell Identification, Formula Writing)
**Platform:** Office.js Excel add-in for Mac (Microsoft 365), TypeScript

## Task List

### Phase 1: Project Foundation

#### Task Group 1: Project Scaffolding & Development Environment
**Dependencies:** None
**Specialist:** DevOps/Build Engineer

- [x] 1.0 Complete project foundation
  - [x] 1.1 Initialize project structure
    - Create root directory structure: `/src/taskpane`, `/src/commands`, `/src/fillgaps`
    - Initialize npm project (`npm init`)
    - Add `.gitignore` for node_modules, dist, and build artifacts
  - [x] 1.2 Install core dependencies
    - Install Office.js: `npm install @microsoft/office-js`
    - Install TypeScript: `npm install --save-dev typescript @types/office-js`
    - Install webpack/bundler: `npm install --save-dev webpack webpack-cli webpack-dev-server`
  - [x] 1.3 Configure TypeScript
    - Create `tsconfig.json` with strict mode enabled
    - Set target: ES2017 minimum for Office.js compatibility
    - Configure output directory and module resolution
    - Enable source maps for debugging
  - [x] 1.4 Set up development tooling
    - Configure webpack for bundling (entry points: taskpane.ts, commands.ts)
    - Set up webpack-dev-server for localhost hosting
    - Add npm scripts: `start`, `build`, `watch`
    - Configure Office Add-in debugging for Mac
  - [x] 1.5 Verify development environment
    - Run `npm run build` successfully
    - Start dev server with `npm start`
    - Confirm localhost serves files (e.g., http://localhost:3000)

**Acceptance Criteria:**
- TypeScript compiles without errors
- Dev server runs and serves files on localhost
- Project structure matches Office.js add-in conventions
- All dependencies install successfully

---

### Phase 2: Manifest & Ribbon Integration

#### Task Group 2: Office.js Manifest Configuration
**Dependencies:** Task Group 1
**Specialist:** Office.js Integration Engineer

- [x] 2.0 Complete manifest and ribbon integration
  - [x] 2.1 Write 2-8 focused tests for manifest validation
    - Limit to 2-8 highly focused tests maximum
    - Test only critical manifest behaviors (e.g., XML validates, ribbon button registers, task pane URL resolves)
    - Skip exhaustive testing of all manifest features
    - Use Office.js manifest validation tools or simple XML schema tests
  - [x] 2.2 Create manifest.xml with basic metadata
    - Use OfficeApp v1.1+ schema
    - Add add-in ID, version (0.1.0), provider name
    - Set display name: "FillGaps"
    - Configure permissions: ReadWriteDocument
    - Reference pattern: Official Office.js manifest samples
  - [x] 2.3 Configure ribbon UI
    - Create ribbon group: "FillGaps"
    - Add button: "Fill Gaps..." with icon
    - Set button action: Open task pane
    - Use standard Office ribbon icons (or placeholder for now)
  - [x] 2.4 Register task pane
    - Set task pane source URL: `https://localhost:3000/taskpane.html`
    - Configure task pane width hint: 300px
    - Set up command bindings for ribbon button
  - [x] 2.5 Sideload add-in in Excel for Mac
    - Copy manifest to Office add-ins folder on Mac
    - Trust self-signed certificate for localhost (if needed)
    - Open Excel, verify "FillGaps" ribbon group appears
    - Click "Fill Gaps..." button, verify task pane attempts to load (may show blank/error for now)
  - [x] 2.6 Ensure manifest tests pass
    - Run ONLY the 2-8 tests written in 2.1
    - Verify manifest validates against schema
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- The 2-8 tests written in 2.1 pass
- Manifest validates against Office.js schema
- Ribbon button appears in Excel for Mac
- Clicking button opens task pane (even if blank)
- No console errors related to manifest loading

---

### Phase 3: Task Pane UI Framework

#### Task Group 3: Task Pane HTML/CSS Structure
**Dependencies:** Task Group 2
**Specialist:** UI Designer/Frontend Engineer

- [x] 3.0 Complete task pane UI framework
  - [x] 3.1 Write 2-8 focused tests for UI rendering
    - Limit to 2-8 highly focused tests maximum
    - Test only critical UI behaviors (e.g., radio buttons render, Run button exists, status area displays text)
    - Skip exhaustive testing of all UI states and interactions
    - Use lightweight DOM testing (e.g., jsdom or manual verification)
  - [x] 3.2 Create taskpane.html structure
    - Add Office.js script reference: `<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>`
    - Add reference to bundled taskpane.js
    - Add reference to taskpane.css
    - Create semantic HTML structure: header, form sections, button, status area
    - File location: `/src/taskpane/taskpane.html`
  - [x] 3.3 Build Target Condition section
    - Section header: "Target Condition"
    - Radio button group (name="targetCondition"):
      - Option 1: "Blanks only" (value="blanks")
      - Option 2: "Errors only" (value="errors")
      - Option 3: "Blanks + Errors" (value="both")
    - Default selection: "Blanks only"
  - [x] 3.4 Build Template Source section
    - Section header: "Template Source"
    - Radio button group (name="templateSource"):
      - Option 1: "Active cell formula" (value="activeCell") - default
      - Option 2: "Top-left cell in selection" (value="topLeft")
    - Default selection: "Active cell formula"
  - [x] 3.5 Add Run button and status area
    - Primary action button: "Run Fill Operation" (id="runButton")
    - Status/feedback text area: `<div id="statusArea"></div>`
    - Ensure button is visually prominent (primary CTA style)
  - [x] 3.6 Create taskpane.css with Office styling
    - File location: `/src/taskpane/taskpane.css`
    - Apply Office Fabric or Fluent UI styling patterns (or minimal custom styling)
    - Ensure 300-400px width compatibility
    - Use clear spacing between sections (margins/padding)
    - Style radio buttons for touch-friendly interaction (Mac trackpad)
    - Style Run button as primary action (e.g., blue background)
    - Style status area with light background for visibility
  - [x] 3.7 Ensure UI rendering tests pass
    - Run ONLY the 2-8 tests written in 3.1
    - Verify critical UI elements render correctly
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- The 2-8 tests written in 3.1 pass
- Task pane displays in Excel with all UI elements visible
- Radio buttons are selectable and mutually exclusive within groups
- Run button is clickable (no action wired yet)
- UI fits within 300-400px width
- Styling follows Office add-in conventions

---

### Phase 4: Task Pane TypeScript Skeleton

#### Task Group 4: Task Pane Event Handlers & Office.js Initialization
**Dependencies:** Task Group 3
**Specialist:** Frontend Engineer

- [x] 4.0 Complete task pane TypeScript skeleton
  - [x] 4.1 Write 2-8 focused tests for Office.js initialization
    - Limit to 2-8 highly focused tests maximum
    - Test only critical initialization behaviors (e.g., Office.initialize runs, form values can be read, Run button handler fires)
    - Skip exhaustive testing of all event handlers
    - Use mocks for Office.js API
  - [x] 4.2 Create taskpane.ts with Office.js initialization
    - File location: `/src/taskpane/taskpane.ts`
    - Implement `Office.initialize = function() { ... }`
    - Wait for Office ready state
    - Attach event listeners after DOM loads
  - [x] 4.3 Implement form value readers
    - Function: `getTargetCondition()` - reads selected radio button from target condition group
    - Function: `getTemplateSource()` - reads selected radio button from template source group
    - Return values: "blanks" | "errors" | "both" and "activeCell" | "topLeft"
  - [x] 4.4 Wire up Run button click handler
    - Add event listener to "runButton"
    - On click, read form values using functions from 4.3
    - Call placeholder function `runFillOperation(targetCondition, templateSource)`
    - Display "Operation started..." in status area
    - Implement try-catch for error handling
  - [x] 4.5 Implement status display helper
    - Function: `updateStatus(message: string, isError: boolean = false)`
    - Updates status area text content
    - Applies error styling if `isError` is true
    - Clears previous status on new operation
  - [x] 4.6 Verify task pane interaction
    - Rebuild project (`npm run build`)
    - Reload add-in in Excel
    - Click Run button, verify "Operation started..." appears in status
    - Change radio button selections, verify values can be read
  - [x] 4.7 Ensure Office.js initialization tests pass
    - Run ONLY the 2-8 tests written in 4.1
    - Verify critical initialization behaviors work
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- The 2-8 tests written in 4.1 pass
- Office.js initializes without errors
- Run button triggers click handler
- Form values can be read correctly
- Status area updates with messages
- No console errors during interaction

---

### Phase 5: Core Engine - Template Detection

#### Task Group 5: Template Formula Detection Module (FR3)
**Dependencies:** Task Group 4
**Specialist:** Backend/API Engineer (Excel.js specialist)

- [x] 5.0 Complete template formula detection
  - [x] 5.1 Write 2-8 focused tests for template detection
    - Limit to 2-8 highly focused tests maximum
    - Test only critical template detection behaviors (e.g., active cell has formula, top-left cell returns R1C1 formula, validation fails on blank cell)
    - Skip exhaustive testing of all edge cases
    - Mock Excel.js API (Range, Cell objects)
  - [x] 5.2 Create fillgaps/types.ts with interfaces
    - File location: `/src/fillgaps/types.ts`
    - Define `IFillOptions` interface: `{ targetCondition: string, templateSource: string }`
    - Define `ITemplateInfo` interface: `{ cell: Excel.Range, formulaR1C1: string }`
    - Define `IFillResult` interface: `{ modifiedCount: number, success: boolean, error?: string }`
  - [x] 5.3 Create fillgaps/engine.ts with template detection
    - File location: `/src/fillgaps/engine.ts`
    - Function: `async getTemplateCell(context: Excel.RequestContext, templateSource: string): Promise<Excel.Range>`
      - If templateSource === "activeCell": return `context.workbook.getActiveCell()`
      - If templateSource === "topLeft": return `context.workbook.getSelectedRange().getCell(0, 0)`
      - Load cell properties before returning
  - [x] 5.4 Implement template formula validation
    - Function: `async validateAndExtractFormula(context: Excel.RequestContext, templateCell: Excel.Range): Promise<string>`
    - Load `templateCell.load("formulasR1C1")`
    - Call `await context.sync()`
    - Validate: formula exists and is not empty string
    - If validation fails, throw error: "Template cell does not contain a formula"
    - Return `templateCell.formulasR1C1[0][0]` (R1C1 formula string)
  - [x] 5.5 Create combined template detection function
    - Function: `async detectTemplateFormula(templateSource: string): Promise<ITemplateInfo>`
    - Wrap logic in `Excel.run(async (context) => { ... })`
    - Call `getTemplateCell()` from 5.3
    - Call `validateAndExtractFormula()` from 5.4
    - Return `{ cell: templateCell, formulaR1C1: formula }`
    - Handle errors gracefully (log to console for now)
  - [x] 5.6 Ensure template detection tests pass
    - Run ONLY the 2-8 tests written in 5.1
    - Verify template detection works for both modes
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- The 2-8 tests written in 5.1 pass
- Template cell can be identified based on user selection
- R1C1 formula can be extracted from template cell
- Validation detects cells without formulas
- Excel.run batch operation executes efficiently
- Errors are logged to console

---

### Phase 6: Core Engine - Cell Identification

#### Task Group 6: Eligible Cell Identification Engine (FR4)
**Dependencies:** Task Group 5
**Specialist:** Backend/API Engineer (Excel.js specialist)

- [x] 6.0 Complete eligible cell identification
  - [x] 6.1 Write 2-8 focused tests for cell eligibility logic
    - Limit to 2-8 highly focused tests maximum
    - Test only critical eligibility behaviors (e.g., blank cell detected, error cell detected, value cell excluded)
    - Skip exhaustive testing of all error types and edge cases
    - Mock Excel.js Range with sample values/formulas
  - [x] 6.2 Implement blank detection helper
    - Function: `isBlank(value: any, formula: any): boolean`
    - Return true if: `value === null || value === undefined || value === ""` AND `formula === "" || formula === null`
    - Truly empty cells only (no formula AND no value)
  - [x] 6.3 Implement error detection helper
    - Function: `isError(value: any): boolean`
    - Check if value is Excel error object or string starts with "#"
    - Detect all error types: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!
    - Use type-based detection, not string matching (Excel.ErrorCodes constants)
  - [x] 6.4 Implement cell eligibility checker
    - Function: `isCellEligible(value: any, formula: any, targetCondition: string): boolean`
    - If targetCondition === "blanks": return `isBlank(value, formula)`
    - If targetCondition === "errors": return `isError(value)`
    - If targetCondition === "both": return `isBlank(value, formula) || isError(value)`
    - Return false for all other cases
  - [x] 6.5 Implement batch cell identification
    - Function: `async identifyEligibleCells(targetCondition: string): Promise<Array<{row: number, col: number}>>`
    - Wrap in `Excel.run(async (context) => { ... })`
    - Get selected range: `context.workbook.getSelectedRange()`
    - Load range properties in single batch: `range.load("values, formulas, rowCount, columnCount")`
    - Call `await context.sync()`
    - Iterate through range cells (nested loop: rows x columns)
    - For each cell, call `isCellEligible()` from 6.4
    - Build array of eligible cell coordinates: `{row: i, col: j}`
    - Return array (empty array if no eligible cells)
  - [x] 6.6 Ensure cell identification tests pass
    - Run ONLY the 2-8 tests written in 6.1
    - Verify blank and error detection work
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- The 2-8 tests written in 6.1 pass
- Blank cells are correctly identified (no formula AND no value)
- Error cells are correctly identified (all error types)
- Non-eligible cells (values, formulas) are excluded
- Batch operation loads range properties efficiently
- Returns empty array when no eligible cells found
- Target condition logic works for all three modes

---

### Phase 7: Core Engine - Formula Writing

#### Task Group 7: Selective Formula Writing Engine (FR5)
**Dependencies:** Task Group 6
**Specialist:** Backend/API Engineer (Excel.js specialist)

- [x] 7.0 Complete selective formula writing
  - [x] 7.1 Write 2-8 focused tests for formula writing
    - Limit to 2-8 highly focused tests maximum
    - Test only critical formula writing behaviors (e.g., formula written to eligible cell, R1C1 reference adjusts correctly, non-eligible cells unchanged)
    - Skip exhaustive testing of all R1C1 reference scenarios
    - Mock Excel.js Range and cell access
  - [x] 7.2 Implement batch formula writing
    - Function: `async writeFormulasToEligibleCells(eligibleCells: Array<{row: number, col: number}>, templateFormula: string): Promise<number>`
    - Wrap in `Excel.run(async (context) => { ... })`
    - Get selected range: `context.workbook.getSelectedRange()`
    - Iterate through `eligibleCells` array
    - For each cell coordinate: `range.getCell(cell.row, cell.col).formulasR1C1 = [[templateFormula]]`
    - Call `await context.sync()` ONCE after all formulas set
    - Return count of modified cells (`eligibleCells.length`)
  - [x] 7.3 Verify R1C1 reference adjustment
    - Test manually in Excel: Create range with formula `=R[-1]C` in A2
    - Use formula writing function to fill A3:A5
    - Verify each cell references the row above (A3 refs A2, A4 refs A3, etc.)
    - Test relative column references: `=RC[-1]`
    - Test absolute references: `=R1C1` (should remain fixed)
    - Test mixed references: `=R[-1]C1` (relative row, absolute column)
  - [x] 7.4 Handle edge case: No eligible cells
    - If `eligibleCells` array is empty, return 0 immediately
    - Do not call Excel.run or context.sync
    - Operation completes successfully with no changes
  - [x] 7.5 Ensure formula writing tests pass
    - Run ONLY the 2-8 tests written in 7.1
    - Verify formulas are written correctly
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- The 2-8 tests written in 7.1 pass
- Formulas are written only to eligible cells
- R1C1 relative references adjust correctly per cell position
- Absolute references remain fixed across filled cells
- Single context.sync call for entire batch
- Returns accurate count of modified cells
- Handles empty eligible cells array gracefully

---

### Phase 8: End-to-End Integration

#### Task Group 8: Integration & Full Pipeline
**Dependencies:** Task Groups 5, 6, 7
**Specialist:** Integration Engineer

- [x] 8.0 Complete end-to-end integration
  - [x] 8.1 Write 2-8 focused tests for full pipeline
    - Limit to 2-8 highly focused tests maximum
    - Test only critical end-to-end workflows (e.g., user clicks Run → blanks filled, user selects errors only → errors filled)
    - Skip exhaustive testing of all configuration combinations
    - Mock Excel.js API for full operation flow
  - [x] 8.2 Implement main fill operation orchestrator
    - Function in `engine.ts`: `async executeFillOperation(options: IFillOptions): Promise<IFillResult>`
    - Step 1: Call `detectTemplateFormula(options.templateSource)`
    - Step 2: Call `identifyEligibleCells(options.targetCondition)`
    - Step 3: If eligible cells found, call `writeFormulasToEligibleCells(eligibleCells, templateFormula)`
    - Step 4: Return `{ modifiedCount, success: true }`
    - Wrap entire pipeline in try-catch, return error result on failure
  - [x] 8.3 Wire orchestrator to task pane
    - Update `runFillOperation()` in `taskpane.ts` from 4.4
    - Import and call `executeFillOperation({ targetCondition, templateSource })`
    - Display result in status area: "Filled {modifiedCount} cells successfully"
    - Handle errors: Display error message in status area with error styling
    - Clear status at start of each operation
  - [x] 8.4 Test end-to-end: Blanks only mode
    - Create test spreadsheet in Excel with:
      - Row 1: Header
      - Row 2: Formula `=A1*2` in A2
      - Rows 3-6: Mix of blanks (A3, A5) and values (A4, A6)
    - Select A2:A6, open FillGaps task pane
    - Set "Blanks only" + "Top-left cell in selection"
    - Click Run
    - Verify: Only A3 and A5 receive formula (A4, A6 unchanged)
    - Verify R1C1 references adjust: A3 = `=A2*2`, A5 = `=A4*2`
  - [x] 8.5 Test end-to-end: Errors only mode
    - Create test spreadsheet with:
      - Column A: Lookup keys (1, 2, 3, 99)
      - Column B: VLOOKUP formulas returning #N/A for row with key 99
    - Select B1:B4, open task pane
    - Set "Errors only" + "Active cell formula" (activate B1 first)
    - Click Run
    - Verify: Only B4 (error cell) receives formula
    - Verify: B1-B3 (existing formulas) unchanged
  - [x] 8.6 Test end-to-end: Blanks + Errors mode
    - Create test spreadsheet with mix:
      - Blank cells: A2, A4
      - Error cells: A3 (contains #DIV/0!)
      - Value cells: A5 (number), A6 (text)
    - Select A2:A6, activate A1 (contains formula)
    - Set "Blanks + Errors" + "Active cell formula"
    - Click Run
    - Verify: A2, A3, A4 receive formula (A5, A6 unchanged)
  - [x] 8.7 Test edge case: No eligible cells
    - Create range A1:A5 with all values (no blanks or errors)
    - Select range, run operation with "Blanks only"
    - Verify: Status shows "Filled 0 cells successfully"
    - Verify: No changes to spreadsheet
  - [x] 8.8 Test edge case: Template has no formula
    - Create range with value (not formula) in A1
    - Select A1:A5, set template source to "Top-left"
    - Click Run
    - Verify: Error message displayed in status area
    - Verify: No changes to spreadsheet
  - [x] 8.9 Test edge case: Active cell outside selection
    - Create formula in A1
    - Select B2:B5 (does not include A1)
    - Set template source to "Active cell formula"
    - Click Run with blanks in B2:B5
    - Verify: B2:B5 receive formula from A1
    - Verify: A1 unchanged
  - [x] 8.10 Ensure end-to-end integration tests pass
    - Run ONLY the 2-8 tests written in 8.1
    - Verify critical workflows work
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- The 2-8 tests written in 8.1 pass
- Full pipeline executes from UI click to formula writing
- All three target condition modes work correctly
- Both template source modes work correctly
- Status area displays operation results
- R1C1 references adjust properly in all scenarios
- Edge cases handled gracefully
- Zero overwrites of existing values or formulas
- Operation completes in < 2 seconds for typical ranges

---

## Testing Philosophy

### Test-Driven Development (TDD) Approach
Each phase follows a focused TDD pattern:
1. **Write 2-8 focused tests** (first sub-task in each group)
2. **Implement feature** (middle sub-tasks)
3. **Run ONLY new tests** (final sub-task in each group)

### Test Scope Limits
- **During Development (Groups 1-7):** Write maximum 2-8 tests per group focusing on critical behaviors only
- **Total Expected Tests:** Approximately 16-56 tests across all groups
- **NO comprehensive testing:** Skip edge cases, performance tests, and exhaustive coverage during development
- **NO full suite runs:** Each group runs only its own tests, not the entire application test suite

### Dedicated Testing Spec
Note: Per requirements document, a comprehensive acceptance test suite will be implemented in a **separate spec (Item 8)**. This spec focuses on minimal strategic testing during development only.

---

## Execution Order

**Recommended implementation sequence:**

1. **Phase 1: Project Foundation** (Task Group 1)
   - Set up development environment before any code

2. **Phase 2: Manifest & Ribbon** (Task Group 2)
   - Establish Office.js integration layer

3. **Phase 3: Task Pane UI** (Task Group 3)
   - Build user interface framework

4. **Phase 4: Task Pane Logic** (Task Group 4)
   - Wire up UI event handlers and Office.js initialization

5. **Phase 5: Template Detection** (Task Group 5)
   - Implement FR3: Template formula detection and validation

6. **Phase 6: Cell Identification** (Task Group 6)
   - Implement FR4: Eligible cell identification engine

7. **Phase 7: Formula Writing** (Task Group 7)
   - Implement FR5: Selective formula writing with R1C1

8. **Phase 8: Integration** (Task Group 8)
   - Connect all components and validate end-to-end workflows

**Parallel Opportunities:**
- Task Groups 3 & 5 can be developed in parallel after Group 2 completes (UI and engine are independent)
- Task Groups 6 & 7 must be sequential (formula writing depends on cell identification)

---

## Key Technical Notes

### Performance Targets
- Handle 1k-10k cells efficiently (< 2 seconds)
- Minimize Excel.run calls (single context per operation)
- Single context.sync per phase (template detection, cell identification, formula writing)
- Target: < 5 total API round trips per operation

### R1C1 Formula Behavior
- **Relative reference:** `=R[-1]C` (cell one row above, adjusts per position)
- **Absolute reference:** `=R1C1` (always references A1, fixed across all cells)
- **Mixed reference:** `=R[-1]C1` (relative row, absolute column)

### Error Types to Detect
All Excel error types must be handled:
- #N/A (notAvailable)
- #VALUE! (value)
- #REF! (reference)
- #DIV/0! (divisionByZero)
- #NUM! (num)
- #NAME? (name)
- #NULL! (null)

### Blank Definition
- **v0 Definition:** Cell has no formula AND no value (truly empty)
- **Not blank:** Cells with formula returning "" (empty string handling deferred to v0.2)

---

## Out of Scope

The following are explicitly NOT included in this spec:

**User-Facing Features:**
- Quick action buttons (Fill Blanks / Fill Errors) - Item 6, separate spec
- Preflight validation with user-facing error dialogs - Item 7, separate spec
- Comprehensive acceptance test suite - Item 8, separate spec
- Empty string as blank toggle - v0.2 feature
- Specific error type filtering (checklist) - v0.2 feature
- Convert to values option - v0.2 feature
- Preview/confirmation dialog - v0.2 feature
- Settings persistence across sessions - v0.2 feature

**Platform Features:**
- Windows platform support - v1 feature
- Context menu integration - v1 feature
- Keyboard shortcuts - v1 feature
- Large range optimization (>10k cells) - v1 feature

**Advanced Scenarios:**
- Multi-area selections (not planned)
- Merged cell special handling (best-effort only)
- Protected sheet detection (Excel API will error naturally)

---

## Success Metrics

### Functional Success
- [x] Add-in loads successfully in Excel for Mac (Microsoft 365)
- [x] Ribbon button "Fill Gaps..." appears and opens task pane
- [x] User can configure all options (3 target conditions × 2 template sources)
- [x] Run button executes fill operation
- [x] Only eligible cells receive formulas (zero overwrites)
- [x] R1C1 references adjust properly across all filled cells

### Performance Success
- [x] Operation completes in < 2 seconds for ranges up to 1k cells
- [x] Single Excel.run context per operation phase
- [x] < 5 total API round trips per complete operation

### Quality Success
- [x] Zero accidental overwrites of existing values or formulas
- [x] All three target condition modes work correctly
- [x] Both template source modes work correctly
- [x] Edge cases handled gracefully (no eligible cells, no template formula, etc.)
- [x] R1C1 relative references adjust correctly 100% of the time

### User Experience Success
- [x] Fill operation completes in 1-2 clicks (vs 10-20 seconds native workflow)
- [x] Task pane UI is responsive and intuitive
- [x] Clear feedback on operation results (status area)
- [x] No console errors during normal operation
