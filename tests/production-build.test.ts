import * as fs from 'fs';
import * as path from 'path';

describe('Production Build Configuration', () => {
  const projectRoot = path.join(__dirname, '..');

  describe('shortcuts.json', () => {
    let shortcutsContent: any;

    beforeAll(() => {
      const shortcutsPath = path.join(projectRoot, 'src/shortcuts.json');
      const content = fs.readFileSync(shortcutsPath, 'utf-8');
      shortcutsContent = JSON.parse(content);
    });

    test('shortcuts.json is valid JSON with correct structure', () => {
      expect(shortcutsContent).toBeDefined();
      expect(shortcutsContent.actions).toBeDefined();
      expect(Array.isArray(shortcutsContent.actions)).toBe(true);
      expect(shortcutsContent.shortcuts).toBeDefined();
      expect(Array.isArray(shortcutsContent.shortcuts)).toBe(true);
    });

    test('shortcuts.json defines fillBlanksCommand and fillErrorsCommand actions', () => {
      const actionIds = shortcutsContent.actions.map((a: any) => a.id);
      expect(actionIds).toContain('fillBlanksCommand');
      expect(actionIds).toContain('fillErrorsCommand');
    });

    test('shortcuts.json defines keyboard shortcuts with Mac keys', () => {
      const fillBlanksShortcut = shortcutsContent.shortcuts.find(
        (s: any) => s.action === 'fillBlanksCommand'
      );
      const fillErrorsShortcut = shortcutsContent.shortcuts.find(
        (s: any) => s.action === 'fillErrorsCommand'
      );

      expect(fillBlanksShortcut).toBeDefined();
      expect(fillBlanksShortcut.key.mac).toBe('Command+Shift+B');
      expect(fillBlanksShortcut.key.default).toBe('Ctrl+Shift+B');

      expect(fillErrorsShortcut).toBeDefined();
      expect(fillErrorsShortcut.key.mac).toBe('Command+Shift+E');
      expect(fillErrorsShortcut.key.default).toBe('Ctrl+Shift+E');
    });
  });

  describe('webpack.config.js', () => {
    test('webpack config exists and exports a function', () => {
      const webpackConfigPath = path.join(projectRoot, 'webpack.config.js');
      expect(fs.existsSync(webpackConfigPath)).toBe(true);

      const webpackConfig = require(webpackConfigPath);
      expect(typeof webpackConfig).toBe('function');
    });
  });

  describe('package.json scripts', () => {
    let packageJson: any;

    beforeAll(() => {
      const packagePath = path.join(projectRoot, 'package.json');
      const content = fs.readFileSync(packagePath, 'utf-8');
      packageJson = JSON.parse(content);
    });

    test('package.json contains production build scripts', () => {
      expect(packageJson.scripts['build:prod']).toBeDefined();
      expect(packageJson.scripts['clean:docs']).toBeDefined();
      expect(packageJson.scripts['deploy']).toBeDefined();
    });

    test('build:prod script includes production environment flag', () => {
      expect(packageJson.scripts['build:prod']).toContain('--env production');
    });
  });
});
