PRD: Fill-Only Blanks / Errors (Excel Add-in, standalone)

1) Product name

FillGaps for Excel (working name)
Tagline: “Fill down, but only where data is missing.”

⸻

2) Problem statement

Excel’s native fill-down workflows (drag fill handle / Fill Down) overwrite everything in the target area. Power users frequently need to fill formulas into a column/range only where values are missing (blank) or where results are invalid (errors), without disturbing existing inputs. There’s a native workaround (Go To Special → Blanks → Ctrl/Cmd+Enter), but it’s multi-step and doesn’t handle “errors only,” “blank-ish,” or one-click conversion to values cleanly.

⸻

3) Target users
	•	Finance / FP&A / Accounting / Ops analysts working with imported datasets and hybrid manual + formula sheets.
	•	Anyone doing reconciliation or augmentation where existing values must not be overwritten.

⸻

4) Goals (v0)

Core
	•	One command to fill a formula into selected range only for cells that are:
	•	blank and/or
	•	error
	•	Does not overwrite non-target cells.
	•	Works for Excel on Mac desktop (Microsoft 365) as an Office.js add-in.
	•	Minimal UI: ribbon buttons + a lightweight settings task pane (or modal).

Quality
	•	Predictable “what is blank” behavior.
	•	Fast for typical ranges (up to ~50k cells is nice, but v0 can target 1k–10k comfortably).

⸻

5) Non-goals (v0)
	•	No custom hotkeys required.
	•	No right-click menu required.
	•	No multi-area selections (single contiguous range only).
	•	No special handling for:
	•	merged cells,
	•	protected sheets that disallow edits,
	•	“intercepting” native Cmd+C/Cmd+V,
	•	complex Table/structured-reference transformations (best-effort only).

⸻

6) Primary use cases

Use case A: Fill blanks only

User has a column with partial data and wants to apply a lookup/calc formula only into empty cells.

Example:
[Red], [blank], [Green], [Blue], [blank] …
They put a lookup formula in the first blank and want the add-in to fill it only into blanks.

Use case B: Fill errors only

User has a calculation column that yields errors for some rows (e.g., missing lookup keys) and wants to apply a fallback formula only where errors exist.

Use case C: Fill blanks + errors

Common “dataset cleanup” mode.

⸻

7) User stories
	1.	As a user, I can select a target range and run Fill Gaps so only blanks get the formula.
	2.	As a user, I can run Fill Gaps so only error cells get the formula.
	3.	As a user, I can choose whether the source formula is taken from:
	•	the active cell (default), or
	•	the top-left cell of the selection (optional toggle).
	4.	As a user, I can optionally convert the newly-filled formulas to values after filling (off by default).

⸻

8) UX specification (v0)

Ribbon commands (minimum viable)

Group: FillGaps
	•	Fill Gaps… (opens small pane/dialog with options + Run)
	•	Fill Blanks (runs with defaults: blanks only, active cell formula)
	•	Fill Errors (runs with defaults: errors only, active cell formula)

Settings (in a small task pane or modal)

Options:
	•	Target condition:
	•	( ) Blanks only
	•	( ) Errors only
	•	( ) Blanks + Errors
	•	Template source:
	•	(•) Active cell formula
	•	( ) Top-left cell in selection
	•	Blank definition (v0 default simple):
	•	(•) Truly empty cells only
	•	( ) Treat "" as blank (v0.2)
	•	Error handling:
	•	(•) Any error
	•	( ) Specific errors (v0.2 checklist)
	•	Post-action:
	•	Convert filled cells to values (v0.2 or v0 optional)

Messaging:
	•	Clear preflight errors, e.g.:
	•	“Active cell must contain a formula.”
	•	“Selection must be a single contiguous range.”
	•	“No blank/error cells found in selection.”

⸻

9) Functional requirements

FR1 — Identify template formula
	•	If “Active cell formula”:
	•	read active cell formula (prefer R1C1 for robust relative fill)
	•	if none: block with message
	•	If “Top-left selection”:
	•	use top-left cell of selected range as the template cell
	•	Store both:
	•	formulaR1C1
	•	formulaA1 (optional fallback)

FR2 — Determine eligible cells in target range

Inputs:
	•	Selected targetRange (must be contiguous)

Eligibility rules:
	•	Blank (v0): cell has no formula and no value (truly empty).
	•	Error (v0): cell’s value is an Excel error of any type.
	•	Blanks + Errors: union.

Important v0 decision:
	•	Cells that display blank because a formula returns "" are NOT considered blank by default.

FR3 — Write into eligible cells only
	•	For each eligible cell:
	•	write templateFormulaR1C1 to that cell.
	•	All other cells remain unchanged.
	•	Best-effort to batch operations for performance (avoid per-cell round trips if possible).

FR4 — Optional convert-to-values (v0.2 or optional v0)
	•	After write, restrict conversion to only the cells the add-in modified.
	•	Replace formulas with their evaluated values.

FR5 — Safety guardrails
	•	If selection size is huge (configurable threshold, e.g. > 200k cells):
	•	warn user and require confirmation (v0.2)
	•	If workbook is in a state that prevents edits:
	•	surface error with actionable instruction.

⸻

10) Edge cases and expected behavior
	•	No eligible cells: no-op + toast (“No blanks/errors found.”)
	•	Active cell is outside selection: allowed (template can be outside selection).
	•	Template references: using R1C1 preserves relative references when filling down/across.
	•	Mixed constants + formulas: only blanks/errors replaced; existing values and formulas preserved.
	•	Filtered range: v0 acts on the selected range regardless of filter visibility.
	•	Tables: works if selection is within; no special structured-ref logic beyond normal formula behavior.

⸻

11) Success metrics (v0)
	•	Time-to-fill gaps reduced from ~10–20 seconds (manual multi-step) to 1–2 clicks.
	•	Zero “unexpected overwrite” incidents in typical workflows.
	•	Reliability: > 95% success on contiguous-range selections up to ~10k cells.

⸻

12) Technical approach (v0)

Platform
	•	Office Add-in (Office.js) for Excel targeting Mac desktop (Microsoft 365).
	•	TypeScript.

Core implementation (high-level)
	•	Excel.run batch:
	1.	Get active cell + selection range.
	2.	Load:
	•	template cell formulasR1C1
	•	target range values and formulas (and error detection)
	3.	Compute eligible cell coordinates.
	4.	Write formulaR1C1 into eligible cells.
	5.	(Optional) convert those cells to values.

Storage
	•	v0: in-memory settings (reset on reload).
	•	v0.2: persist settings using Office storage so defaults stick per user.

Packaging
	•	v0: sideloaded manifest on your Mac for personal use.
	•	v1: signed distribution approach (AppSource vs direct), plus licensing.

⸻

13) Roadmap

v0 (personal Mac utility)
	•	Ribbon commands + minimal settings
	•	Blanks only / Errors only / Both
	•	Template from active cell
	•	Works reliably on contiguous ranges

v0.2 (polish)
	•	Treat "" as blank toggle
	•	Select specific error types
	•	Convert filled to values option
	•	“Preview: N cells will be modified” confirmation

v1 (sellable Mac + Windows)
	•	Add right-click context menu entry
	•	Add configurable keyboard shortcut
	•	Licensing + telemetry (optional, privacy-respecting)

⸻

14) Acceptance tests (must-pass)
	1.	Blanks only

	•	Given a column with values in rows 2,3,5 and blanks in 4,6
	•	And active cell contains a formula
	•	When running Fill Blanks on rows 2–6
	•	Then only rows 4 and 6 receive the formula.

	2.	Errors only

	•	Given a column with #N/A in rows 10–12 and valid values elsewhere
	•	When running Fill Errors
	•	Then only rows 10–12 receive the formula.

	3.	Do not overwrite

	•	Given a range containing existing manual values and formulas
	•	When running any FillGaps mode
	•	Then all non-eligible cells remain byte-for-byte unchanged.

	4.	Template required

	•	If template source cell contains no formula
	•	Command fails with message and makes no edits.