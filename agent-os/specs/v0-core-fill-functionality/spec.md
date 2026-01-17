# Specification: FillGaps Core Fill Functionality (v0)

## Goal
Deliver a working Office.js Excel add-in for Mac (Microsoft 365) that fills formulas into selected ranges only where cells are blank and/or contain errors, without overwriting existing values or formulas. This walking skeleton provides immediate user value by automating the tedious multi-step Go To Special workflow.

## User Stories
- As a Finance analyst, I want to select a range and fill formulas only into blank cells so that I preserve all existing manual inputs and formulas
- As an Operations analyst, I want to fill fallback formulas only into cells with errors (#N/A, #VALUE!, etc.) so that I can quickly fix lookup failures without manual cell-by-cell editing
- As a data professional, I want to choose whether my template formula comes from the active cell or the top-left of my selection so that I have flexibility in my workflow
- As a user filling gaps, I want clear visual feedback when the operation completes so that I know which cells were modified

## Core Requirements
- User can select a contiguous range in Excel and open the FillGaps task pane from the ribbon
- User can configure target condition (Blanks only / Errors only / Blanks + Errors)
- User can choose template formula source (Active cell / Top-left cell in selection)
- User can click Run to execute the fill operation
- System fills formulas using R1C1 notation only into eligible cells based on target condition
- System preserves all non-eligible cells unchanged (values and formulas)
- System handles edge cases gracefully (no eligible cells, template has no formula, etc.)

## Visual Design
No visual mockups provided for v0. UI will follow Office add-in conventions:

**Ribbon UI:**
- Group name: "FillGaps"
- Button label: "Fill Gaps..." (opens task pane)
- Uses standard Office ribbon icons and styling

**Task Pane Layout:**
- Header: "FillGaps Settings"
- Section 1: Target Condition (radio buttons)
  - Blanks only
  - Errors only
  - Blanks + Errors
- Section 2: Template Source (radio buttons)
  - Active cell formula (default)
  - Top-left cell in selection
- Run button (primary action)
- Status/feedback text area (for operation results)
- Standard Office task pane width (300-400px)

**Responsive Requirements:**
- Task pane must work at standard Office add-in widths (300-400px)
- Controls must be touch-friendly for Mac trackpad gestures

## Reusable Components

### Existing Code to Leverage
This is a greenfield project with no existing codebase. However, the implementation will follow established Office.js patterns:

**Office.js Standard Patterns:**
- Excel.run() batch operations for efficient API calls
- Context-based loading of range properties (values, formulas, formulasR1C1)
- Standard manifest XML structure for ribbon customization
- Task pane HTML/CSS/TypeScript pattern

**Excel.js API Features:**
- Range.formulasR1C1 for robust relative reference handling
- Range.values for reading cell contents
- Range.getUsedRange() patterns for efficient cell access
- Error type detection using cell value inspection

### New Components Required

**FR1: Manifest Configuration (New)**
- XML manifest defining add-in metadata, ribbon group, and task pane
- Ribbon command registration for "Fill Gaps..." button
- Permissions declaration for workbook read/write access
- WHY: Office add-ins require manifest for Excel integration

**FR2: Task Pane UI (New)**
- HTML form with radio button groups for target condition and template source
- Run button with click handler
- Status display area for operation feedback
- CSS styling following Office Fabric / Fluent UI conventions
- WHY: Users need visual interface to configure fill options

**FR3: Template Detection Module (New)**
- Function to determine template source cell based on user selection
- Formula validation logic to ensure template contains a formula
- R1C1 formula extraction from template cell
- WHY: Core requirement to identify and validate source formula

**FR4: Cell Eligibility Engine (New)**
- Batch operation to load range values and formulas
- Blank detection logic (cell has no formula AND no value)
- Error detection logic (cell value is Excel error type)
- Eligible cell coordinate list builder
- WHY: Core requirement to identify which cells should receive formula

**FR5: Formula Writing Engine (New)**
- Batch operation to write R1C1 formulas to eligible cells only
- Cell-by-cell formula assignment within single Excel.run context
- Sync operation to commit changes to workbook
- WHY: Core requirement to selectively fill formulas without overwriting

## Technical Approach

**Platform & Architecture:**
- Office.js add-in (TypeScript) targeting Excel for Mac (Microsoft 365)
- Browser-based runtime (Office.js runs in embedded browser within Excel)
- Manifest XML (OfficeApp v1.1+) for add-in registration
- Task pane HTML + TypeScript for UI
- Excel.js API for all workbook interactions

**File Structure:**
```
/manifest.xml                 # Add-in manifest defining ribbon and task pane
/src
  /taskpane
    /taskpane.html           # Task pane UI markup
    /taskpane.ts             # Task pane UI logic and event handlers
    /taskpane.css            # Task pane styling
  /commands
    /commands.ts             # Ribbon command handlers
  /fillgaps
    /engine.ts               # Core fill logic (FR3, FR4, FR5)
    /types.ts                # TypeScript interfaces and types
/package.json                # npm dependencies (Office.js, TypeScript)
/tsconfig.json              # TypeScript compiler configuration
```

**Core Algorithm (engine.ts):**
```
1. Validate user has selected a range
2. Determine template source cell:
   - If "Active cell": use context.workbook.getActiveCell()
   - If "Top-left selection": use range.getCell(0, 0)
3. Load template cell formulaR1C1 property
4. Validate template has formula (not blank, not just value)
5. Get selected range and load values + formulas properties
6. Iterate through range cells:
   - For each cell, check eligibility based on target condition:
     * Blanks only: cell.value === null AND cell.formula === ""
     * Errors only: cell.value is error type (#N/A, #VALUE!, etc.)
     * Blanks + Errors: either of above
7. For each eligible cell, set cell.formulaR1C1 = template.formulaR1C1
8. Call context.sync() to commit batch changes
9. Return count of modified cells
```

**Key APIs:**
- `Excel.run(async (context) => { ... })` for all operations
- `context.workbook.getSelectedRange()` to get target range
- `context.workbook.getActiveCell()` for active cell template source
- `range.load("values, formulas, formulasR1C1")` for efficient property loading
- `range.getCell(rowIndex, colIndex)` for cell-specific access
- `cell.formulaR1C1 = templateFormula` for formula assignment
- `context.sync()` to commit batched changes

**Error Detection:**
Excel error types to detect:
- `#N/A` (Excel.ErrorCodes.notAvailable)
- `#VALUE!` (Excel.ErrorCodes.value)
- `#REF!` (Excel.ErrorCodes.reference)
- `#DIV/0!` (Excel.ErrorCodes.divisionByZero)
- `#NUM!` (Excel.ErrorCodes.num)
- `#NAME?` (Excel.ErrorCodes.name)
- `#NULL!` (Excel.ErrorCodes.null)

Implementation: Check if cell.value type is error object or string starts with "#"

**R1C1 vs A1:**
- Use R1C1 notation exclusively for formula template and writing
- R1C1 preserves relative references correctly when filling across rows/columns
- Example: `=R[-1]C` always references cell one row above, regardless of position

**Batch Performance:**
- Single Excel.run() context for entire operation
- Load all range properties in one batch
- Write all formulas in same batch before sync
- Target: Handle 1k-10k cells efficiently (< 2 seconds)

## Out of Scope

**Features NOT in this spec (future specs):**
- Quick action buttons (Fill Blanks / Fill Errors) - Separate spec for Item 6
- Preflight validation with user-facing error messages - Separate spec for Item 7
- Acceptance test suite - Separate spec for Item 8
- Treat empty string ("") as blank toggle - v0.2 feature
- Specific error type filtering (checklist) - v0.2 feature
- Convert to values option - v0.2 feature
- Preview/confirmation dialog - v0.2 feature
- Settings persistence across sessions - v0.2 feature
- Context menu integration - v1 feature
- Keyboard shortcuts - v1 feature
- Windows platform support - v1 feature
- Large range optimization (>10k cells) - v1 feature

**Technical Constraints:**
- Multi-area selections not supported (single contiguous range only)
- Merged cells not specially handled (best-effort only)
- Protected sheets will fail with Excel API error
- Filtered ranges processed regardless of filter visibility
- Tables/structured references work with normal formula behavior (no special handling)

## Edge Cases & Behaviors

**No Eligible Cells Found:**
- Operation completes successfully with no changes
- Future spec will add user feedback ("No blank/error cells found")
- Current spec: Operation succeeds silently

**Template Cell Has No Formula:**
- Must be detected in FR3 validation
- Future spec will add error UI
- Current spec: Operation fails (error logged to console)

**Active Cell Outside Selection:**
- Allowed when using "Active cell formula" mode
- Template can be anywhere in workbook
- Validation: Template cell must exist and contain formula

**Template Cell Inside Selection:**
- Allowed in both modes
- If template cell is eligible, it will be overwritten with its own formula (no-op)

**Mixed Formulas and Values in Range:**
- All existing formulas preserved unchanged
- All existing values preserved unchanged
- Only truly blank or error cells modified

**R1C1 Reference Behavior:**
- Relative references (e.g., `=R[-1]C`) adjust correctly per cell position
- Absolute references (e.g., `=R1C1`) remain fixed across all filled cells
- Mixed references (e.g., `=R[-1]C1`) combine relative and absolute correctly

**Selection Edge Cases:**
- Single cell selection: Allowed (may have 0 or 1 eligible cells)
- Entire column selection: Allowed (will process all used cells)
- Large range (>10k cells): May be slow; no warning in v0 (v0.2 will add threshold)

**Error Type Variations:**
- v0 treats all error types equally (any error matches "Errors only" mode)
- Error strings may vary by Excel locale (e.g., "#N/A" vs "#N/V")
- Detection should be type-based, not string-based

**Formula Syntax Preservation:**
- Template formula syntax preserved exactly (including whitespace)
- R1C1 notation ensures references adjust automatically
- Named ranges and structured references pass through unchanged

## Success Criteria

**Functional Success:**
- Add-in loads successfully in Excel for Mac (Microsoft 365)
- Ribbon button "Fill Gaps..." appears in FillGaps group
- Task pane opens when button clicked
- User can select target condition (3 options)
- User can select template source (2 options)
- Run button triggers fill operation
- Only eligible cells receive formulas
- Non-eligible cells remain unchanged
- R1C1 references adjust properly across filled cells

**Quality Metrics:**
- Zero accidental overwrites of existing values or formulas
- Formula filling completes in < 2 seconds for ranges up to 1k cells
- Operation succeeds reliably on valid inputs (95%+ success rate)
- R1C1 relative references adjust correctly 100% of the time

**User Experience:**
- User can complete fill operation in 1-2 clicks (vs 10-20 seconds with native workflow)
- Task pane UI is responsive and intuitive
- Radio button selections persist during session (in-memory only for v0)

**Technical Success:**
- TypeScript compiles without errors
- Manifest validates against Office.js schema
- Add-in can be sideloaded on Mac without errors
- Excel.run batch operations execute efficiently (< 5 API round trips per operation)

## Dependencies

**Required for Development:**
- Office.js library (latest stable version)
- TypeScript compiler (v4.5+)
- Node.js and npm for package management
- Excel for Mac (Microsoft 365) for testing and debugging
- Office Add-in Debugger (built-in to Excel for Mac)

**Optional Development Tools:**
- ESLint + Prettier for code quality
- Webpack or Rollup for production bundling
- Office Yeoman Generator (yo office) for project scaffolding

**Runtime Dependencies:**
- Excel for Mac (Microsoft 365) - target platform
- Office.js runtime (provided by Excel host)
- Modern browser engine (provided by Office.js embedded runtime)

**Technical Prerequisites:**
- Mac computer with Microsoft 365 Excel installed
- Developer mode enabled in Excel for add-in sideloading
- Localhost web server for hosting add-in files during development

## Implementation Notes

**Standards Compliance:**
- Follow Office.js best practices for batch operations (minimize sync calls)
- Use TypeScript strict mode for type safety
- Implement fail-fast validation (check preconditions early)
- Provide user-friendly error messages (future spec will add UI)
- Keep functions small and focused (single responsibility)
- Use descriptive variable names (no abbreviations)
- Remove dead code and commented blocks

**Naming Conventions:**
- Files: lowercase with hyphens (task-pane.html, fill-engine.ts)
- TypeScript interfaces: PascalCase with 'I' prefix (IFillOptions)
- Functions: camelCase describing action (detectTemplateFormula)
- Constants: UPPER_SNAKE_CASE (DEFAULT_TARGET_CONDITION)
- CSS classes: lowercase with hyphens (fill-gaps-button)

**Error Handling:**
- Centralize error handling in task pane event handlers
- Use try-catch around all Excel.run() calls
- Log errors to console for debugging (v0)
- Future spec will add user-facing error messages
- Fail explicitly with clear error messages
- Clean up resources in finally blocks

**Testing Approach (for future spec):**
- Focus on core user flows only
- Test behavior, not implementation
- Mock Excel.js API for unit tests
- Manual acceptance testing in Excel for Mac
- Defer edge case testing to dedicated phase

## Future Enhancements

**v0.2 Features (Not in this spec):**
- Empty string handling toggle
- Specific error type filtering
- Convert to values option
- Preview/confirmation dialog
- Settings persistence using Office storage API

**v1 Features (Not in this spec):**
- Quick action buttons (Fill Blanks, Fill Errors)
- Context menu integration
- Keyboard shortcut configuration
- Windows platform support
- Large range optimization with progress indication
- Licensing and distribution via AppSource

**Potential Improvements:**
- Multi-area selection support
- Merged cell handling
- Protected sheet detection with helpful errors
- Undo/redo integration
- Batch operation progress feedback
- Custom blank definition (formula returns "")
