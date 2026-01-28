# Task Breakdown: Production Deployment with Hotkeys and Context Menu

## Overview
Total Tasks: 4 Task Groups with 21 sub-tasks

This spec enables the FillGaps Excel add-in to run on multiple machines without a development server by deploying to GitHub Pages, and adds keyboard shortcuts plus context menu integration for faster access to fill operations.

## Task List

### Build & Deployment Infrastructure

#### Task Group 1: Production Build Configuration
**Dependencies:** None

- [x] 1.0 Complete production build infrastructure
  - [x] 1.1 Write 2-4 focused tests for production build output
    - Test that `npm run build:prod` generates required files in `/docs` folder
    - Test that shortcuts.json is included in build output
    - Test that asset paths are correctly configured
  - [x] 1.2 Update webpack.config.js for production deployment
    - Add environment variable or CLI flag to output to `/docs` folder for production
    - Ensure all assets (HTML, JS, CSS, icons) are copied correctly
    - Add shortcuts.json to CopyWebpackPlugin patterns
    - Reference pattern: existing webpack.config.js structure at `/Users/michaelchristopher/repos/excel_conditional_fill/webpack.config.js`
  - [x] 1.3 Update package.json with production build scripts
    - Add `build:prod` script that outputs to `/docs` folder
    - Add `deploy` script for GitHub Pages deployment workflow
  - [x] 1.4 Create shortcuts.json file
    - Location: `/Users/michaelchristopher/repos/excel_conditional_fill/src/shortcuts.json`
    - Define keyboard shortcut mappings for fillBlanksCommand and fillErrorsCommand
    - Mac shortcuts: `Command+Shift+B` (Fill Blanks), `Command+Shift+E` (Fill Errors)
    - Windows shortcuts: `Ctrl+Shift+B` (Fill Blanks), `Ctrl+Shift+E` (Fill Errors)
  - [x] 1.5 Create deploy.sh script
    - Location: `/Users/michaelchristopher/repos/excel_conditional_fill/scripts/deploy.sh`
    - Automate: clean docs folder, run production build, copy production manifest
    - Include verification steps for required files
  - [x] 1.6 Verify production build tests pass
    - Run only the 2-4 tests written in 1.1
    - Verify `/docs` folder contains all required files
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- `npm run build:prod` generates files in `/docs` folder
- All required files present: taskpane.html, commands.html, taskpane.js, commands.js, shortcuts.json, assets/
- shortcuts.json follows Office.js keyboard shortcut schema
- deploy.sh script runs successfully

### Manifest Configuration

#### Task Group 2: Production Manifest and Extensions
**Dependencies:** Task Group 1

- [x] 2.0 Complete manifest configuration for production
  - [x] 2.1 Write 2-4 focused tests for manifest validation
    - Test that manifest-production.xml is valid XML
    - Test that all URLs point to GitHub Pages
    - Test that ContextMenu ExtensionPoint is properly structured
    - Test that ExtendedOverrides references shortcuts.json
  - [x] 2.2 Create manifest-production.xml
    - Location: `/Users/michaelchristopher/repos/excel_conditional_fill/manifest-production.xml`
    - Copy from existing manifest.xml
    - Replace all `https://localhost:3000/` URLs with `https://mikedtcm22.github.io/excel_conditional_fill/`
    - Reference pattern: existing manifest.xml at `/Users/michaelchristopher/repos/excel_conditional_fill/manifest.xml`
  - [x] 2.3 Add ContextMenu ExtensionPoint to manifest-production.xml
    - Add `<ExtensionPoint xsi:type="ContextMenu">` within DesktopFormFactor
    - Add OfficeMenu id="ContextMenuCell" with Fill Blanks and Fill Errors controls
    - Reuse existing resource strings (FillBlanksButton.Label, FillErrorsButton.Label, etc.)
    - Each control invokes existing fillBlanksCommand/fillErrorsCommand via ExecuteFunction
  - [x] 2.4 Add ExtendedOverrides for keyboard shortcuts to manifest-production.xml
    - Add ExtendedOverrides element pointing to shortcuts.json URL on GitHub Pages
    - URL: `https://mikedtcm22.github.io/excel_conditional_fill/shortcuts.json`
  - [x] 2.5 Update development manifest.xml with context menu
    - Add same ContextMenu ExtensionPoint to development manifest
    - Add ExtendedOverrides pointing to localhost shortcuts.json
    - Ensures feature parity between dev and production
  - [x] 2.6 Verify manifest tests pass
    - Run only the 2-4 tests written in 2.1
    - Validate both manifest files with XML schema
    - Do NOT run the entire test suite at this stage

**Acceptance Criteria:**
- manifest-production.xml validates without XML errors
- All URLs in production manifest point to GitHub Pages
- ContextMenu entries appear with correct labels and actions
- ExtendedOverrides references shortcuts.json correctly
- Development manifest also updated with context menu support

### GitHub Pages Deployment

#### Task Group 3: GitHub Pages Configuration and Documentation
**Dependencies:** Task Groups 1, 2

- [x] 3.0 Complete GitHub Pages deployment and documentation
  - [x] 3.1 Run production build to populate /docs folder
    - Execute `npm run build:prod`
    - Verify all files generated in `/docs`
    - Copy manifest-production.xml to `/docs` folder
  - [x] 3.2 Configure GitHub repository for GitHub Pages
    - Enable GitHub Pages in repository settings
    - Set source to `main` branch, `/docs` folder
    - Verify HTTPS is enabled (automatic with GitHub Pages)
    - Note: This is a manual step in GitHub repository settings
  - [x] 3.3 Create PRODUCTION_INSTALL.md documentation
    - Location: `/Users/michaelchristopher/repos/excel_conditional_fill/PRODUCTION_INSTALL.md`
    - Step-by-step sideloading instructions for Mac
    - Include: download manifest, create wef folder, copy manifest, restart Excel
    - Document keyboard shortcuts (Cmd+Shift+B, Cmd+Shift+E)
    - Document context menu usage
    - Include troubleshooting section
  - [ ] 3.4 Verify GitHub Pages deployment
    - Navigate to `https://mikedtcm22.github.io/excel_conditional_fill/`
    - Verify taskpane.html loads
    - Verify commands.html loads
    - Verify shortcuts.json loads with correct CORS headers
    - Verify assets (icons) load correctly
    - Note: Requires GitHub Pages to be enabled and deployment pushed

**Acceptance Criteria:**
- GitHub Pages serves all files at `https://mikedtcm22.github.io/excel_conditional_fill/`
- HTTPS certificate is valid
- shortcuts.json accessible with correct MIME type
- PRODUCTION_INSTALL.md contains complete sideloading instructions
- Documentation includes keyboard shortcut reference

### Verification & Testing

#### Task Group 4: Manual Testing and Verification
**Dependencies:** Task Groups 1-3

- [ ] 4.0 Complete end-to-end verification
  - [ ] 4.1 Review existing tests and identify critical gaps
    - Review tests written in Task Groups 1-2
    - Focus on integration between components
    - Total existing tests: approximately 4-8 tests
  - [ ] 4.2 Manual testing: Production deployment
    - Verify files served from GitHub Pages with correct MIME types
    - Verify manifest-production.xml downloads correctly
    - Verify HTTPS works without certificate warnings
  - [ ] 4.3 Manual testing: Sideload on primary machine
    - Follow PRODUCTION_INSTALL.md instructions
    - Copy manifest to `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`
    - Restart Excel and verify add-in appears in My Add-ins
    - Verify add-in loads without dev server
  - [ ] 4.4 Manual testing: Keyboard shortcuts
    - Test `Cmd+Shift+B` triggers Fill Blanks on Mac
    - Test `Cmd+Shift+E` triggers Fill Errors on Mac
    - Verify shortcuts do not conflict with Excel native shortcuts
    - Verify shortcuts work when task pane is closed
  - [ ] 4.5 Manual testing: Context menu
    - Right-click on selected cells
    - Verify "Fill Blanks" appears in context menu
    - Verify "Fill Errors" appears in context menu
    - Test both menu entries trigger correct commands
    - Verify error dialogs appear for invalid operations
  - [ ] 4.6 Manual testing: Cross-access method consistency
    - Select same range and execute Fill Blanks via ribbon, keyboard, and context menu
    - Verify all three methods produce identical results
    - Select same range and execute Fill Errors via ribbon, keyboard, and context menu
    - Verify all three methods produce identical results
  - [ ] 4.7 Optional: Test on second machine
    - If second machine available, follow PRODUCTION_INSTALL.md
    - Verify fresh install works without dev server
    - Verify all access methods work

**Acceptance Criteria:**
- Add-in loads and functions from GitHub Pages without local server
- Keyboard shortcuts `Cmd+Shift+B` and `Cmd+Shift+E` work on Mac
- Context menu entries appear and function when right-clicking cells
- All three access methods (ribbon, keyboard, context menu) produce identical results
- Manual testing checklist from spec completed

## Execution Order

Recommended implementation sequence:

1. **Task Group 1: Production Build Configuration** - Sets up the build infrastructure required for all other tasks
2. **Task Group 2: Production Manifest and Extensions** - Creates manifest files with context menu and keyboard shortcuts
3. **Task Group 3: GitHub Pages Deployment** - Deploys to GitHub Pages and creates user documentation
4. **Task Group 4: Manual Testing and Verification** - End-to-end verification of all features

## Notes

### Dependencies Between Task Groups
- Task Group 2 requires shortcuts.json from Task Group 1
- Task Group 3 requires build output from Task Group 1 and manifests from Task Group 2
- Task Group 4 requires all previous groups to be complete

### Key File Locations
- Production manifest: `/Users/michaelchristopher/repos/excel_conditional_fill/manifest-production.xml`
- Shortcuts configuration: `/Users/michaelchristopher/repos/excel_conditional_fill/src/shortcuts.json`
- Deploy script: `/Users/michaelchristopher/repos/excel_conditional_fill/scripts/deploy.sh`
- Installation docs: `/Users/michaelchristopher/repos/excel_conditional_fill/PRODUCTION_INSTALL.md`
- Build output: `/Users/michaelchristopher/repos/excel_conditional_fill/docs/`

### Existing Code to Leverage
- Command handlers: `/Users/michaelchristopher/repos/excel_conditional_fill/src/commands/commands.ts` (no changes needed)
- Fill engine: `/Users/michaelchristopher/repos/excel_conditional_fill/src/fillgaps/engine.ts` (no changes needed)
- Validation: `/Users/michaelchristopher/repos/excel_conditional_fill/src/fillgaps/validation.ts` (no changes needed)
- Webpack config: `/Users/michaelchristopher/repos/excel_conditional_fill/webpack.config.js` (modify for production output)
- Current manifest: `/Users/michaelchristopher/repos/excel_conditional_fill/manifest.xml` (reference for production manifest)

### Testing Approach
This spec focuses on deployment infrastructure and manifest configuration. Per the spec, automated testing is limited since:
- Manifest validation is best done via XML schema validation
- Shortcut registration requires runtime Office.js environment
- Context menu appearance requires manual verification in Excel
- Cross-machine deployment requires actual second machine

Total automated tests: approximately 4-8 focused tests across Task Groups 1-2
Primary verification: Manual testing checklist in Task Group 4
