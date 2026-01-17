# FillGaps v0 Core Fill Functionality - Specification

## Overview
This specification covers the core "walking skeleton" functionality for FillGaps Excel add-in (Items 1-5 from the roadmap). This represents the minimal viable product that delivers real user value: selectively filling formulas into blank and/or error cells without overwriting existing data.

## Spec Documents

### Main Specification
**File:** `spec.md`

Comprehensive specification covering:
- User stories and core requirements
- Visual design (ribbon UI and task pane)
- Functional requirements (FR1-FR5)
- Technical implementation approach
- Edge cases and behaviors
- Success criteria
- Dependencies and constraints

**Key Sections:**
- FR1: Manifest & Ribbon Setup
- FR2: Task Pane UI
- FR3: Template Formula Detection
- FR4: Eligible Cell Identification
- FR5: Selective Formula Writing

### Planning Documents

**File:** `planning/requirements.md`

Requirements analysis based on source documents:
- PRD analysis and scope definition
- Roadmap items covered (1-5)
- Functional requirements breakdown
- Technical requirements and APIs
- Edge cases and validated behaviors
- Success criteria and acceptance tests

## Scope

### In Scope (Items 1-5)
1. Manifest & Ribbon UI - Office.js add-in with ribbon button
2. Settings Task Pane - UI with target condition and template source options
3. Template Formula Detection - Validate and read R1C1 formula from source
4. Eligible Cell Identification - Batch identify blank/error cells
5. Selective Formula Writing - Batch write formulas only to eligible cells

### Out of Scope (Future Specs)
- Item 6: Quick Action Commands (Fill Blanks / Fill Errors buttons)
- Item 7: Validation & Error Messages (user-facing)
- Item 8: End-to-End Testing (acceptance test suite)
- v0.2 features: Empty string handling, specific error filtering, convert to values
- v1 features: Context menu, keyboard shortcuts, Windows support

## Implementation Guidance

### File Structure
```
/manifest.xml                 # Add-in manifest
/src
  /taskpane
    /taskpane.html           # Task pane UI
    /taskpane.ts             # UI logic and event handlers
    /taskpane.css            # Styling
  /commands
    /commands.ts             # Ribbon command handlers
  /fillgaps
    /engine.ts               # Core fill logic (FR3-FR5)
    /types.ts                # TypeScript interfaces
/package.json                # Dependencies
/tsconfig.json              # TypeScript config
```

### Key Technical Decisions
- **R1C1 formulas:** Use R1C1 notation exclusively for robust relative reference handling
- **Batch operations:** Single Excel.run() context per operation for performance
- **Blank definition:** v0 treats only truly empty cells as blank (no formula AND no value)
- **Error detection:** All Excel error types treated equally (#N/A, #VALUE!, etc.)
- **Contiguous range only:** Single selection required; multi-area not supported

### Success Metrics
- Zero accidental overwrites of existing data
- Operation completes in < 2 seconds for 1k cells
- Reduces workflow from 10-20 seconds to 1-2 clicks
- R1C1 references adjust correctly 100% of the time

## Platform
- **Target:** Excel for Mac (Microsoft 365)
- **Framework:** Office.js Add-in (TypeScript)
- **Runtime:** Browser-based (embedded in Excel)
- **Development:** Node.js + npm + TypeScript compiler

## Next Steps
1. Review this specification for completeness
2. Set up development environment (Office.js + TypeScript)
3. Implement FR1-FR5 in sequence
4. Manual testing in Excel for Mac
5. Proceed to Item 6 spec (Quick Action Commands)
