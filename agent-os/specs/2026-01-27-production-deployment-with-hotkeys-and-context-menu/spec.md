# Specification: Production Deployment with Hotkeys and Context Menu

## Goal

Enable the FillGaps Excel add-in to run on multiple machines without a development server by deploying to GitHub Pages, and add keyboard shortcuts plus context menu integration for faster access to fill operations.

## User Stories

- As a user, I want to install the add-in on my work and home computers without running a dev server so that I can use FillGaps anywhere
- As a user, I want to trigger Fill Blanks with a keyboard shortcut so that I can work faster without using the ribbon
- As a user, I want to trigger Fill Errors with a keyboard shortcut so that I can work faster without using the ribbon
- As a user, I want to right-click and select fill operations from a context menu so that I have quick access when my hands are on the mouse

## Core Requirements

- Build production-ready static assets (HTML, JS, CSS, icons) optimized for deployment
- Deploy static assets to GitHub Pages with HTTPS support
- Create a production manifest.xml referencing GitHub Pages URLs instead of localhost
- Add Mac-friendly keyboard shortcuts for Fill Blanks and Fill Errors
- Add context menu entries for Fill Blanks and Fill Errors when cells are selected
- Provide sideloading instructions for Mac users

## Visual Design

No visual mockups provided. This spec focuses on deployment infrastructure and non-visual UI access methods (shortcuts and context menus).

## Reusable Components

### Existing Code to Leverage

- **Commands**: `/Users/michaelchristopher/repos/excel_conditional_fill/src/commands/commands.ts`
  - `fillBlanksCommand` and `fillErrorsCommand` handlers already implemented
  - Uses `Office.actions.associate()` for command registration
  - Follows try/catch with `showErrorDialog` pattern

- **Engine**: `/Users/michaelchristopher/repos/excel_conditional_fill/src/fillgaps/engine.ts`
  - `executeFillOperation` function - no changes needed

- **Validation**: `/Users/michaelchristopher/repos/excel_conditional_fill/src/fillgaps/validation.ts`
  - `showErrorDialog` and validation functions - no changes needed

- **Manifest**: `/Users/michaelchristopher/repos/excel_conditional_fill/manifest.xml`
  - Existing ribbon button definitions and resource string patterns
  - VersionOverrides V1.0 structure in place

- **Webpack**: `/Users/michaelchristopher/repos/excel_conditional_fill/webpack.config.js`
  - Production build already configured (`npm run build`)
  - Outputs to `/dist` directory

- **Assets**: `/Users/michaelchristopher/repos/excel_conditional_fill/assets/`
  - Icons already exist (icon-16.png, icon-32.png, icon-64.png, icon-80.png)

### New Components Required

- **Production Manifest**: `manifest-production.xml` - Separate manifest with GitHub Pages URLs
- **Keyboard Shortcuts JSON**: `src/shortcuts.json` - Defines keyboard shortcut mappings
- **Deploy Script**: `scripts/deploy.sh` - Automates GitHub Pages deployment
- **GitHub Actions Workflow** (optional): `.github/workflows/deploy.yml` - Automated deployment on push

## Technical Approach

### 1. Production Deployment to GitHub Pages

**GitHub Pages URL Pattern:**
- Repository: `mikedtcm22/excel_conditional_fill`
- GitHub Pages URL: `https://mikedtcm22.github.io/excel_conditional_fill/`

**Deployment Strategy:**
- Use `/docs` folder as GitHub Pages source (simpler than gh-pages branch)
- Build assets to `/docs` folder instead of `/dist` for production
- Configure GitHub Pages in repository settings to deploy from `/docs` on `main` branch

**Webpack Configuration Updates:**
- Add production output path option for `/docs` folder
- Ensure all assets (HTML, JS, CSS, icons) are copied correctly

**Production Manifest:**
- Create `manifest-production.xml` with all URLs pointing to GitHub Pages
- Replace all `https://localhost:3000/` references with `https://mikedtcm22.github.io/excel_conditional_fill/`

### 2. Keyboard Shortcuts (Mac-Friendly)

**Research Findings:**
Office.js keyboard shortcuts on Mac support these modifiers:
- `Command` (Cmd) - Primary command modifier on Mac
- `Option` - Mac equivalent of Alt
- `Shift` - Must be combined with Cmd or Option
- `Control` - Available but less common on Mac

**Recommended Shortcuts:**
Based on Office.js best practices for Mac compatibility:

| Action | Mac Shortcut | Rationale |
|--------|--------------|-----------|
| Fill Blanks | `Command+Shift+B` | Cmd+Shift is standard Mac pattern, B for Blanks |
| Fill Errors | `Command+Shift+E` | Cmd+Shift is standard Mac pattern, E for Errors |

**Implementation via shortcuts.json:**
```json
{
  "actions": [
    {
      "id": "fillBlanksCommand",
      "type": "ExecuteFunction",
      "name": "Fill Blanks"
    },
    {
      "id": "fillErrorsCommand",
      "type": "ExecuteFunction",
      "name": "Fill Errors"
    }
  ],
  "shortcuts": [
    {
      "action": "fillBlanksCommand",
      "key": {
        "default": "Ctrl+Shift+B",
        "mac": "Command+Shift+B"
      }
    },
    {
      "action": "fillErrorsCommand",
      "key": {
        "default": "Ctrl+Shift+E",
        "mac": "Command+Shift+E"
      }
    }
  ]
}
```

**Manifest Integration:**
- Add `ExtendedOverrides` element pointing to shortcuts.json URL
- shortcuts.json must be served from GitHub Pages with proper CORS headers

### 3. Context Menu Integration

**ExtensionPoint Type:** `ContextMenu` with `OfficeMenu id="ContextMenuCell"`

**Menu Structure:**
Add two entries to the cell context menu:
- "Fill Blanks" - Triggers `fillBlanksCommand`
- "Fill Errors" - Triggers `fillErrorsCommand`

**Manifest XML Addition:**
```xml
<ExtensionPoint xsi:type="ContextMenu">
  <OfficeMenu id="ContextMenuCell">
    <Control xsi:type="Button" id="ContextMenuFillBlanks">
      <Label resid="FillBlanksButton.Label"/>
      <Supertip>
        <Title resid="FillBlanksButton.Title"/>
        <Description resid="FillBlanksButton.Desc"/>
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="Icon.16x16"/>
        <bt:Image size="32" resid="Icon.32x32"/>
        <bt:Image size="80" resid="Icon.80x80"/>
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>fillBlanksCommand</FunctionName>
      </Action>
    </Control>
    <Control xsi:type="Button" id="ContextMenuFillErrors">
      <Label resid="FillErrorsButton.Label"/>
      <Supertip>
        <Title resid="FillErrorsButton.Title"/>
        <Description resid="FillErrorsButton.Desc"/>
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="Icon.16x16"/>
        <bt:Image size="32" resid="Icon.32x32"/>
        <bt:Image size="80" resid="Icon.80x80"/>
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>fillErrorsCommand</FunctionName>
      </Action>
    </Control>
  </OfficeMenu>
</ExtensionPoint>
```

## File Changes Required

### New Files

| File | Purpose |
|------|---------|
| `manifest-production.xml` | Production manifest with GitHub Pages URLs |
| `src/shortcuts.json` | Keyboard shortcut definitions |
| `scripts/deploy.sh` | Build and deploy script |
| `PRODUCTION_INSTALL.md` | Installation instructions for production use |

### Modified Files

| File | Changes |
|------|---------|
| `manifest.xml` | Add ContextMenu ExtensionPoint, add ExtendedOverrides for shortcuts |
| `webpack.config.js` | Add production output to `/docs` folder option |
| `package.json` | Add deploy scripts |

### Generated Files (Build Output)

| Location | Contents |
|----------|----------|
| `/docs/` | Production build output for GitHub Pages |
| `/docs/taskpane.html` | Task pane HTML |
| `/docs/commands.html` | Commands HTML |
| `/docs/taskpane.js` | Bundled task pane JS |
| `/docs/commands.js` | Bundled commands JS |
| `/docs/shortcuts.json` | Keyboard shortcuts config |
| `/docs/assets/` | Icon files |

## User Workflow

### One-Time Setup (Each Machine)

1. Download `manifest-production.xml` from GitHub repository
2. Create sideload folder if it does not exist:
   ```bash
   mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef
   ```
3. Copy manifest to sideload folder:
   ```bash
   cp manifest-production.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/FillGaps-manifest.xml
   ```
4. Restart Excel
5. Insert add-in via Insert > My Add-ins > Developer Add-ins > FillGaps

### Daily Usage

1. Open Excel workbook
2. Select a range (active cell contains template formula)
3. Trigger fill operation via:
   - **Keyboard**: `Cmd+Shift+B` (Fill Blanks) or `Cmd+Shift+E` (Fill Errors)
   - **Context Menu**: Right-click > Fill Blanks / Fill Errors
   - **Ribbon**: Click Fill Blanks or Fill Errors button
4. Eligible cells are filled with the template formula

## Testing Approach

### Manual Testing Checklist

**Deployment:**
- [ ] `npm run build:prod` creates files in `/docs` folder
- [ ] All required files present (HTML, JS, assets, shortcuts.json)
- [ ] GitHub Pages serves files with correct MIME types
- [ ] HTTPS certificate is valid for GitHub Pages URL

**Manifest:**
- [ ] Production manifest validates (no XML errors)
- [ ] All URLs resolve to correct GitHub Pages paths
- [ ] ExtendedOverrides URL is correct

**Keyboard Shortcuts:**
- [ ] `Cmd+Shift+B` triggers Fill Blanks on Mac
- [ ] `Cmd+Shift+E` triggers Fill Errors on Mac
- [ ] Shortcuts do not conflict with Excel native shortcuts
- [ ] Shortcuts work when task pane is closed

**Context Menu:**
- [ ] "Fill Blanks" appears in cell right-click menu
- [ ] "Fill Errors" appears in cell right-click menu
- [ ] Context menu entries trigger correct commands
- [ ] Error dialogs appear for invalid operations

**End-to-End:**
- [ ] Fresh install on second machine works without dev server
- [ ] All three access methods (keyboard, context menu, ribbon) produce identical results
- [ ] Error handling works correctly for all access methods

### Automated Testing

Existing tests cover core functionality. No new automated tests required for this spec since it focuses on deployment and manifest configuration. Manual testing is appropriate for:
- Manifest validation
- Shortcut registration
- Context menu appearance
- Cross-machine deployment

## Out of Scope

- **Windows platform support** - Mac only for this spec
- **AppSource submission** - Personal sideloading only
- **Custom build CI/CD** - Manual deployment acceptable
- **Settings persistence** - Planned for v0.2
- **Empty string as blank option** - Planned for v0.2
- **Specific error type filtering** - Planned for v0.2
- **Convert to values option** - Planned for v0.2
- **Preview and confirmation** - Planned for v0.2
- **Large range optimization** - Planned for v1
- **Licensing system** - Planned for v1
- **Telemetry** - Planned for v1

## Success Criteria

- Add-in loads and functions correctly from GitHub Pages without any local server
- Keyboard shortcuts `Cmd+Shift+B` and `Cmd+Shift+E` successfully trigger fill operations on Mac
- Context menu entries appear and function when right-clicking cells
- Second machine can install and use the add-in by following documented instructions
- All three access methods (ribbon, keyboard, context menu) produce identical fill results
