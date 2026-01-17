# Tech Stack

## Platform & Runtime
- **Platform:** Office.js Add-in for Excel
- **Target Application:** Excel for Mac desktop (Microsoft 365) — v0 initial target
- **Language:** TypeScript
- **Runtime:** Browser-based (Office.js runs in embedded browser runtime within Excel)
- **Package Manager:** npm

## Core APIs & Libraries
- **Excel.js API:** Primary API for Excel object model interaction
- **Office.js:** Core Office add-in framework for host communication
- **Excel.run Batch Operations:** Context-based batching for efficient API calls
- **formulasR1C1:** R1C1 reference style for robust relative formula handling
- **Range API:** Cell range selection, value reading, and formula writing

## UI Components
- **Ribbon Commands:** Custom ribbon group with command buttons registered via manifest
- **Task Pane:** HTML/CSS/JS-based side panel for settings and configuration
- **Office UI Fabric / Fluent UI:** (Optional) Microsoft's design system for Office add-ins
- **Vanilla JavaScript/TypeScript:** Lightweight UI interactions without heavy framework

## Add-in Architecture
- **Manifest:** XML manifest defining add-in metadata, ribbon customizations, and permissions
- **Commands:** Ribbon button handlers invoking TypeScript functions
- **Settings Storage:** Office storage API for persisting user preferences (v0.2+)
- **Error Handling:** Try-catch with user-facing error messages via Office dialog or task pane

## Development & Build Tools
- **TypeScript Compiler:** Type-safe development with Excel.js type definitions
- **Webpack / Rollup:** (Optional) Module bundling for production builds
- **Office Add-in Debugger:** Built-in Office developer tools for debugging on Mac
- **Linting/Formatting:** ESLint + Prettier for code quality

## Testing & Quality
- **Test Framework:** Jest or Mocha for unit testing TypeScript logic
- **Office.js Testing:** Manual acceptance testing in Excel for Mac (Microsoft 365)
- **Type Checking:** TypeScript strict mode for compile-time validation
- **Edge Case Testing:** Acceptance test suite covering blanks-only, errors-only, no overwrites, template validation

## Deployment & Distribution
- **v0 Packaging:** Sideloaded XML manifest for personal use on Mac
- **v0.2 Distribution:** Direct installation via shared manifest file
- **v1 Distribution Options:**
  - **AppSource:** Microsoft's official add-in marketplace (requires validation and approval)
  - **Direct Distribution:** Signed manifest distributed outside AppSource
- **Hosting:** Static web hosting for add-in HTML/JS/CSS files (GitHub Pages, Azure, or CDN)
- **Versioning:** Semantic versioning in manifest and package.json

## Key Technical Decisions
- **R1C1 vs A1 Formulas:** R1C1 chosen for robust relative reference handling when filling across rows/columns
- **Batch vs Per-Cell Operations:** Excel.run batch operations for performance (avoid round trips)
- **Contiguous Range Only:** v0 simplification—no multi-area selections to reduce complexity
- **No Merged Cell Handling:** v0 scope—merged cells not specially handled
- **Formula-Only Template:** Template must contain formula (not static value) to prevent user error

## Platform Expansion (v1)
- **Windows Support:** Extend compatibility to Excel for Windows (Microsoft 365)
- **Cross-Platform Testing:** Verify Office.js API parity between Mac and Windows
- **Platform-Specific Quirks:** Address any Mac vs Windows behavioral differences in Excel.js API

## Third-Party Services
- **Licensing (v1):** TBD—potential integration with license key validation service
- **Telemetry (v1):** Privacy-respecting analytics (opt-in) for usage patterns and error tracking
- **Monitoring:** (Optional) Sentry or similar for error reporting in production
