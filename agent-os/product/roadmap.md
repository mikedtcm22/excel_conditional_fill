# Product Roadmap

## v0: Personal Mac Utility (MVP)

1. [x] **Manifest & Ribbon UI** — Create Office.js add-in manifest with ribbon group containing "Fill Gaps...", "Fill Blanks", and "Fill Errors" buttons that register with Excel for Mac (Microsoft 365). `S`

2. [x] **Settings Task Pane** — Build lightweight task pane with radio buttons for target condition (Blanks only / Errors only / Blanks + Errors), template source selection (Active cell / Top-left), and Run button with clear action messaging. `S`

3. [x] **Template Formula Detection** — Implement logic to read and validate formula from active cell or top-left selection, extract formulaR1C1 for robust relative reference handling, and block execution with clear error if no formula exists. `S`

4. [x] **Eligible Cell Identification** — Build Excel.run batch operation that loads target range values and formulas, identifies truly empty cells (blank mode) and/or cells with any error type (error mode), and computes coordinate list of eligible cells. `M`

5. [x] **Selective Formula Writing** — Implement batch write operation using R1C1 formulas to fill only eligible cells while preserving all non-target cells unchanged, with proper error handling for edge cases like no eligible cells found. `M`

6. [ ] **Quick Action Commands** — Wire up "Fill Blanks" and "Fill Errors" ribbon buttons to execute with sensible defaults (active cell template, respective target modes) for one-click workflows. `S`

7. [ ] **Validation & Error Messages** — Add preflight checks for contiguous range selection, formula presence in template cell, and provide clear user-facing error messages for invalid states. `S`

8. [ ] **End-to-End Testing** — Create acceptance tests covering: blanks-only fill preserves values/formulas, errors-only fill targets only error cells, no overwrites occur, template required validation. `M`

## v0.2: Polish & Enhancement

9. [ ] **Treat Empty String as Blank** — Add toggle to blank definition settings allowing users to treat formula-generated "" as blank cells eligible for filling. `S`

10. [ ] **Specific Error Type Filtering** — Expand error handling settings with checklist to target specific error types (#N/A, #VALUE!, #REF!, #DIV/0!, etc.) rather than any error. `S`

11. [ ] **Convert to Values Option** — Add post-action toggle to convert newly-filled formulas to static values, tracking which cells were modified to avoid converting unrelated cells. `M`

12. [ ] **Preview & Confirmation** — Implement preflight preview showing "N cells will be modified" with confirmation step before execution, especially for large ranges. `S`

13. [ ] **Settings Persistence** — Use Office storage API to persist user's default settings (target mode, blank definition, template source preference) across Excel sessions. `S`

## v1: Sellable Product (Mac + Windows)

14. [ ] **Context Menu Integration** — Add right-click context menu entry for "Fill Gaps..." when range is selected, providing quick access without ribbon navigation. `M`

15. [ ] **Keyboard Shortcut Configuration** — Implement configurable keyboard shortcuts for Fill Blanks and Fill Errors quick actions with user preference storage. `M`

16. [ ] **Windows Platform Support** — Test and adapt add-in for Excel on Windows (Microsoft 365), ensuring cross-platform Office.js API compatibility and addressing any platform-specific quirks. `L`

17. [ ] **Large Range Optimization** — Enhance batch operation performance for ranges up to 50k cells with progress indication for operations exceeding configurable threshold (~10k cells). `M`

18. [ ] **Licensing & Distribution** — Implement licensing system (one-time purchase or subscription model) and prepare for AppSource distribution or direct signed distribution approach. `L`

19. [ ] **Privacy-Respecting Telemetry** — Add optional anonymous usage analytics (feature usage frequency, range sizes, error rates) with clear opt-in/opt-out and privacy policy compliance. `M`

> Notes
> - Roadmap progresses from v0 (personal Mac utility) → v0.2 (polish) → v1 (commercial product)
> - v0 focuses on core selective fill functionality with minimal UI for personal use
> - v0.2 adds convenience features and configuration persistence
> - v1 expands to Windows, adds monetization, and includes power-user features
> - Each item represents end-to-end (frontend + backend) functional and testable feature
> - Technical dependencies: Items 1-2 must complete before 3-8; Items 3-5 are the core engine; Item 8 validates all v0 functionality
