# Spec Requirements: Production Deployment with Hotkeys and Context Menu

## Initial Description

For the next spec, something I'd like to build out is the possibility for this to be usable on a different machine (i.e. be downloaded as an add-in & installed to run - rather than currently requiring I manually start up a dev server to run it). Beyond that, I think the main features I'd like to implement would be hotkey and right-click-menu access. The remaining features suggested in the roadmap beyond that are all good, but I think they can wait to be upgrades once I have a stable version of the current functionality (plus hotkey/right-click) that I can install on a 2nd computer.

## Requirements Discussion

### First Round Questions

**Q1:** Deployment approach - personal use on multiple machines vs. broad distribution?
**Answer:** Personal use only (work laptop + home computer).

**Q2:** Hosting solution for static files?
**Answer:** GitHub Pages for free static hosting.

**Q3:** Installation method?
**Answer:** Sideloading via manifest file.

**Q4:** Which access methods to implement - keyboard shortcuts, context menu, or both?
**Answer:** Both. Keyboard shortcuts: Alt+B (Fill Blanks), Alt+E (Fill Errors). Context menu: "Fill Blanks" and "Fill Errors" entries.

**Q5:** User workflow confirmation?
**Answer:** Confirmed:
1. User selects range (active cell provides template formula)
2. User triggers via hotkey OR right-click context menu
3. Fill operation runs on eligible cells in selection

**Q6:** Platform support?
**Answer:** Mac/Excel for Mac support (personal use). Windows support excluded from this spec.

### Existing Code to Reference

**Similar Features Identified:**
- Feature: Ribbon Commands - Path: `/Users/michaelchristopher/repos/excel_conditional_fill/src/commands/commands.ts`
  - Contains existing `fillBlanksCommand` and `fillErrorsCommand` handlers
  - Uses `executeFillOperation` from engine with `templateSource: 'activeCell'` and appropriate `targetCondition`
  - Follows pattern of try/catch with `showErrorDialog` for errors and `event.completed()` for command completion
- Feature: Manifest Configuration - Path: `/Users/michaelchristopher/repos/excel_conditional_fill/manifest.xml`
  - Current ribbon button definitions in `ExtensionPoint xsi:type="PrimaryCommandSurface"`
  - Resource strings pattern for labels and descriptions
  - Currently uses `https://localhost:3000/` URLs that need updating to GitHub Pages URLs

## Visual Assets

### Files Provided:
No visual assets provided.

### Visual Insights:
N/A

## Requirements Summary

### Functional Requirements

**1. Production Deployment to GitHub Pages**
- Build production-ready static assets (HTML, JS, CSS, icons)
- Deploy to GitHub Pages with HTTPS support
- Update manifest.xml to reference GitHub Pages URLs instead of localhost
- Create production manifest file for sideloading

**2. Keyboard Shortcuts**
- Alt+B: Trigger Fill Blanks command
- Alt+E: Trigger Fill Errors command
- Shortcuts must invoke the existing `fillBlanksCommand` and `fillErrorsCommand` handlers
- Must work on Excel for Mac (Microsoft 365)

**3. Context Menu Integration**
- Add "Fill Blanks" entry to right-click context menu when range is selected
- Add "Fill Errors" entry to right-click context menu when range is selected
- Context menu entries must invoke the existing command handlers
- Menu entries should appear in a "FillGaps" submenu or as individual entries

**4. User Workflow**
- User selects a range in Excel (active cell contains the template formula)
- User triggers fill via:
  - Keyboard shortcut (Alt+B or Alt+E), OR
  - Right-click context menu ("Fill Blanks" or "Fill Errors"), OR
  - Existing ribbon buttons (unchanged)
- Fill operation executes on eligible cells within selection
- Error handling via existing `showErrorDialog` pattern

### Reusability Opportunities

**Existing Code to Leverage:**
- `src/commands/commands.ts` - Command handlers already implemented, keyboard shortcuts and context menu should invoke these same functions
- `src/fillgaps/engine.ts` - Core fill engine, no changes needed
- `src/fillgaps/validation.ts` - Validation and error display, no changes needed
- `manifest.xml` - Base structure exists, needs extension for shortcuts and context menu, URL updates for production

**Patterns to Follow:**
- Office.actions.associate() pattern for command registration
- ExecuteFunction action type for UI-less commands
- Resource strings pattern (bt:ShortStrings, bt:LongStrings) for labels

### Scope Boundaries

**In Scope:**
- Production build configuration and deployment scripts
- GitHub Pages deployment (static hosting)
- Production manifest.xml with GitHub Pages URLs
- Keyboard shortcut definitions (Alt+B, Alt+E) in manifest
- Context menu entries in manifest (ContextMenu ExtensionPoint)
- Sideloading instructions for Mac
- Mac/Excel for Mac support

**Out of Scope:**
- Windows platform support
- AppSource submission or distribution
- Settings persistence (v0.2 feature)
- Treat empty string as blank (v0.2 feature)
- Specific error type filtering (v0.2 feature)
- Convert to values option (v0.2 feature)
- Preview & confirmation (v0.2 feature)
- Large range optimization (v1 feature)
- Licensing system (v1 feature)
- Telemetry (v1 feature)

### Technical Considerations

**Manifest Updates Required:**
- Add `Keyboard` ExtensionPoint with shortcut definitions
- Add `ContextMenu` ExtensionPoint for right-click menu entries
- Update all URL resources from `https://localhost:3000/` to GitHub Pages URL
- May need VersionOverrides 1.1 for keyboard shortcuts support

**Build & Deployment:**
- Webpack production build configuration
- Asset optimization (minification, bundling)
- GitHub Pages deployment (likely via `gh-pages` branch or `/docs` folder)
- HTTPS is required for Office add-ins (GitHub Pages provides this)

**Keyboard Shortcut Constraints:**
- Office.js keyboard shortcuts may have platform-specific behavior
- Alt+key combinations are standard for Office add-in shortcuts
- Must not conflict with Excel's native shortcuts

**Context Menu Constraints:**
- ContextMenu ExtensionPoint adds entries to cell right-click menu
- Entries appear when cells are selected
- Uses same ExecuteFunction action type as ribbon buttons

**Sideloading on Mac:**
- Manifest file placed in `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
- Or loaded via Insert > Add-ins > My Add-ins > Upload My Add-in
- No code signing required for personal sideloading
