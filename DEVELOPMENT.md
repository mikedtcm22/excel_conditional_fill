# FillGaps Development Guide

## Project Structure

```
excel_conditional_fill/
├── src/
│   ├── taskpane/          # Task pane UI components
│   │   ├── taskpane.html
│   │   ├── taskpane.ts
│   │   └── taskpane.css   (to be created)
│   ├── commands/          # Ribbon command handlers
│   │   └── commands.ts
│   └── fillgaps/          # Core fill logic engine
│       ├── engine.ts      (to be created)
│       └── types.ts       (to be created)
├── dist/                  # Build output (generated)
├── manifest.xml           # Office Add-in manifest
├── package.json           # npm dependencies and scripts
├── tsconfig.json          # TypeScript configuration
└── webpack.config.js      # Webpack bundler configuration
```

## Development Environment

### Prerequisites
- Node.js (v14+)
- npm
- Excel for Mac (Microsoft 365)

### Installation
```bash
npm install
```

### Available Scripts

- `npm run build` - Production build
- `npm start` - Start development server with hot reload
- `npm run watch` - Watch mode for development

### Development Server
The dev server runs on `https://localhost:3000` and serves the add-in files for Office.js integration.

### Build Process
- TypeScript files are compiled from `src/` to `dist/`
- Webpack bundles the application with source maps enabled in development mode
- Entry points: `taskpane.ts` and `commands.ts`

## Technology Stack

- **Runtime:** Office.js (Excel Add-in API)
- **Language:** TypeScript (strict mode, ES2017 target)
- **Bundler:** Webpack 5
- **Dev Server:** webpack-dev-server (HTTPS)

## Next Steps

Refer to `agent-os/specs/v0-core-fill-functionality/tasks.md` for implementation tasks.
