# Requirements: FillGaps Core Fill Functionality (v0)

## Source Documents
- PRD: `/Users/michaelchristopher/repos/excel_conditional_fill/agent-os/excel-fill-only-PRD.md`
- Product Mission: `/Users/michaelchristopher/repos/excel_conditional_fill/agent-os/product/mission.md`
- Roadmap: `/Users/michaelchristopher/repos/excel_conditional_fill/agent-os/product/roadmap.md`
- Tech Stack: `/Users/michaelchristopher/repos/excel_conditional_fill/agent-os/product/tech-stack.md`

## Roadmap Scope
This spec covers v0 Core functionality (Items 1-5 from roadmap):

1. **Manifest & Ribbon UI** - Office.js add-in manifest with ribbon group containing buttons
2. **Settings Task Pane** - Lightweight task pane UI with target condition options and Run button
3. **Template Formula Detection** - Read and validate formula from active cell or top-left selection (R1C1)
4. **Eligible Cell Identification** - Batch operation to identify blank/error cells in target range
5. **Selective Formula Writing** - Batch write R1C1 formulas only to eligible cells

## Problem Statement
Excel's native fill-down workflows overwrite everything in the target area. Finance, FP&A, Accounting, and Operations analysts need to fill formulas into ranges only where values are missing (blank) or invalid (errors), without disturbing existing inputs. The native workaround (Go To Special → Blanks → Ctrl/Cmd+Enter) is multi-step, doesn't handle errors, and takes 10-20 seconds per operation.

## User Goals
- Fill formulas into blank cells only, preserving all existing data
- Fill fallback formulas into error cells only, fixing lookup failures quickly
- Choose template formula source (active cell or top-left selection)
- Execute fill operation in 1-2 clicks instead of 10-20 seconds
- Zero risk of accidental overwrites

## Functional Requirements

### FR1: Manifest & Ribbon UI
- XML manifest (OfficeApp v1.1+) defining add-in metadata
- Ribbon group named "FillGaps"
- Ribbon button "Fill Gaps..." that opens task pane
- Permissions for workbook read/write access
- Command registration for Excel on Mac (Microsoft 365)

### FR2: Settings Task Pane
- HTML task pane with form controls
- Target condition radio buttons:
  - Blanks only
  - Errors only
  - Blanks + Errors
- Template source radio buttons:
  - Active cell formula (default)
  - Top-left cell in selection
- Run button (primary action)
- Status/feedback text area
- Standard Office add-in styling (300-400px width)

### FR3: Template Formula Detection
- Determine template source cell based on user selection
- Read `formulaR1C1` property from template cell
- Validate that cell contains a formula (not blank, not just value)
- Store template formula for application to eligible cells
- Handle edge case: active cell outside selection (allowed)

### FR4: Eligible Cell Identification
- Load target range (must be contiguous selection)
- Load range values and formulas in single batch
- For each cell in range, determine eligibility:
  - **Blanks only**: Cell has no formula AND no value (truly empty)
  - **Errors only**: Cell value is Excel error (#N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!)
  - **Blanks + Errors**: Union of above conditions
- Build list of eligible cell coordinates
- Handle edge case: no eligible cells (operation completes with no changes)

### FR5: Selective Formula Writing
- Use Excel.run() batch operation for efficiency
- For each eligible cell, set `cell.formulaR1C1 = template.formulaR1C1`
- Preserve all non-eligible cells unchanged
- Call `context.sync()` to commit batch changes
- Use R1C1 notation for proper relative reference adjustment
- Return count of modified cells

## Technical Requirements

### Platform
- Office.js Excel add-in (TypeScript)
- Target: Excel for Mac desktop (Microsoft 365)
- Browser-based runtime (Office.js embedded browser)

### Architecture
- Manifest XML for add-in registration
- Task pane HTML/CSS/TypeScript for UI
- Core engine module for fill logic
- Excel.js API for all workbook interactions

### Key APIs
- `Excel.run()` for batch operations
- `context.workbook.getSelectedRange()`
- `context.workbook.getActiveCell()`
- `range.load("values, formulas, formulasR1C1")`
- `range.getCell(rowIndex, colIndex)`
- `cell.formulaR1C1` for formula assignment
- `context.sync()` for committing changes

### Performance Targets
- Handle 1k-10k cells efficiently (< 2 seconds)
- Single Excel.run() context per operation
- Minimize sync calls (< 5 API round trips)

## Edge Cases

### Validated Behaviors
- **No eligible cells**: Operation completes successfully with no changes
- **Template has no formula**: Validation fails (future spec adds UI)
- **Active cell outside selection**: Allowed in "Active cell formula" mode
- **Template inside selection**: Allowed; may overwrite itself (no-op)
- **Mixed formulas and values**: All existing content preserved unchanged
- **R1C1 relative references**: Adjust correctly per cell position
- **Single cell selection**: Allowed (may have 0 or 1 eligible cells)
- **Large ranges**: May be slow; no warning in v0

### Error Detection
- All Excel error types treated equally in v0
- Detect by cell value type, not string matching
- Error types: #N/A, #VALUE!, #REF!, #DIV/0!, #NUM!, #NAME?, #NULL!

### Blank Definition
- v0: Truly empty cells only (no formula AND no value)
- Cells with formula returning "" NOT considered blank
- Future v0.2: Toggle to treat "" as blank

## Out of Scope for This Spec
- Quick action buttons (Fill Blanks / Fill Errors) - Item 6, separate spec
- Preflight validation with user-facing error messages - Item 7, separate spec
- Acceptance test suite - Item 8, separate spec
- Empty string as blank toggle - v0.2 feature
- Specific error type filtering - v0.2 feature
- Convert to values option - v0.2 feature
- Preview/confirmation dialog - v0.2 feature
- Settings persistence - v0.2 feature
- Context menu integration - v1 feature
- Keyboard shortcuts - v1 feature
- Windows platform support - v1 feature
- Multi-area selections - Not planned
- Merged cell handling - Best-effort only
- Protected sheet detection - Excel API will error naturally

## Success Criteria
- Add-in loads successfully in Excel for Mac
- User can open task pane from ribbon button
- User can configure target condition and template source
- Clicking Run fills only eligible cells with correct formula
- Non-eligible cells remain unchanged (zero overwrites)
- R1C1 references adjust properly across filled cells
- Operation completes in < 2 seconds for typical ranges (1k cells)
- Reduces fill operation time from 10-20 seconds to 1-2 clicks

## Dependencies
- Office.js library (latest stable)
- TypeScript compiler (v4.5+)
- Node.js and npm
- Excel for Mac (Microsoft 365)
- Office Add-in Debugger (built-in)

## User Stories Addressed
From PRD section 7:
1. ✓ User can select a range and run Fill Gaps so only blanks get the formula
2. ✓ User can run Fill Gaps so only error cells get the formula
3. ✓ User can choose whether source formula is from active cell or top-left of selection
4. ✗ Convert to values after filling (v0.2 feature, not in this spec)

## Acceptance Tests (from PRD)
These will be implemented in separate Item 8 spec:
1. Blanks only: Given column with values in rows 2,3,5 and blanks in 4,6, when running Fill Blanks on rows 2-6, then only rows 4,6 receive formula
2. Errors only: Given column with #N/A in rows 10-12, when running Fill Errors, then only rows 10-12 receive formula
3. Do not overwrite: Given range with existing values and formulas, when running any FillGaps mode, then all non-eligible cells remain unchanged
4. Template required: If template cell contains no formula, command fails with message and makes no edits
