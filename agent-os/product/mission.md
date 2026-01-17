# Product Mission

## Pitch
FillGaps for Excel is a smart fill add-in that helps Finance, FP&A, Accounting, and Operations analysts fill formulas into partial datasets by selectively targeting only blank cells and/or error cells, without overwriting existing values or formulas.

## Users

### Primary Customers
- **Finance & FP&A Analysts**: Professionals managing budgets, forecasts, and financial models who work with hybrid datasets mixing manual inputs and formulas
- **Accounting Teams**: Accountants performing reconciliation and data augmentation across imported spreadsheets
- **Operations Analysts**: Data professionals cleaning up imported datasets and maintaining hybrid manual + formula worksheets

### User Personas
**Sarah, Financial Analyst** (28-45)
- **Role:** FP&A Analyst at mid-sized company
- **Context:** Works with monthly budget reports combining imported actuals and manual forecasts
- **Pain Points:** Spends 10-20 seconds per fill operation using Go To Special → Blanks → Ctrl+Enter workflow; native fill-down overwrites critical manual inputs; no easy way to fill only error cells with fallback formulas
- **Goals:** Reduce data prep time, eliminate risk of accidentally overwriting existing values, handle missing lookup values efficiently

**David, Operations Analyst** (25-40)
- **Role:** Operations Data Analyst managing inventory and logistics reports
- **Context:** Imports partial datasets from multiple systems and needs to apply calculations only where data is missing
- **Pain Points:** Multi-step workarounds don't handle errors; must manually identify and fix cells with #N/A or other errors; risk of breaking existing formulas when filling
- **Goals:** One-click solution to fill gaps in data, handle both blanks and errors in single operation, maintain data integrity

## The Problem

### Excel's Fill-Down Overwrites Everything
Excel's native fill-down workflows (drag fill handle or Fill Down command) overwrite all cells in the target area—including existing values, formulas, and manual inputs. Power users frequently need to apply formulas only where data is missing (blank cells) or invalid (error cells), without disturbing existing entries.

The native workaround (Go To Special → Blanks → Ctrl/Cmd+Enter) requires multiple steps, doesn't handle "errors only" or "blanks + errors" modes, and provides no clean way to convert results to values. Users waste 10-20 seconds per operation and risk making mistakes.

**Our Solution:** One-click ribbon commands that intelligently fill formulas into only the cells you specify (blanks, errors, or both), preserving all existing data and formulas. Reduce fill operations from 10-20 seconds to 1-2 clicks.

## Differentiators

### Selective Fill Without Overwriting
Unlike Excel's native fill-down which overwrites everything, we provide surgical precision filling that targets only blank and/or error cells. This results in zero accidental overwrites and complete preservation of existing values and formulas.

### Purpose-Built for Hybrid Datasets
Unlike complex macro solutions or VBA scripts, FillGaps is a lightweight Office.js add-in designed specifically for the modern analyst workflow—imported data + formulas + manual inputs. No programming required, works natively with Excel for Mac (Microsoft 365).

### Three Modes for Common Workflows
Unlike the single-purpose Go To Special workaround, we provide three dedicated modes:
- **Fill Blanks**: Apply formulas only to empty cells
- **Fill Errors**: Apply fallback formulas only to cells with errors (#N/A, #VALUE!, etc.)
- **Fill Blanks + Errors**: Dataset cleanup mode for comprehensive gap-filling

This eliminates the need for multiple manual operations and reduces cognitive load.

## Key Features

### Core Features
- **Selective Fill Modes**: Choose to fill only blank cells, only error cells, or both—without touching existing values or formulas
- **Formula Template Source**: Use either the active cell or top-left selection as your formula template, automatically adapted using R1C1 relative references for robust filling
- **Ribbon Quick Actions**: One-click "Fill Blanks" and "Fill Errors" buttons for instant execution with sensible defaults

### Configuration Features
- **Target Condition Settings**: Radio button selection for Blanks only, Errors only, or Blanks + Errors combined
- **Blank Definition Control**: Define whether truly empty cells count as blank (v0), with option to treat formula-generated "" as blank (v0.2)
- **Error Type Filtering**: Handle any error type by default (v0), with ability to target specific error types like #N/A or #VALUE! (v0.2)

### Advanced Features
- **Convert to Values**: Optionally convert newly-filled formula results to static values after filling (v0.2)
- **Preflight Validation**: Clear error messages for invalid states ("Active cell must contain a formula", "No blank/error cells found")
- **Batch Performance**: Efficient R1C1-based batch operations handle typical ranges (1k-10k cells) quickly and reliably
