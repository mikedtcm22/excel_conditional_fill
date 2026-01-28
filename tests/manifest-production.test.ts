import * as fs from 'fs';
import * as path from 'path';
import { parseString } from 'xml2js';
import { promisify } from 'util';

const parseXml = promisify(parseString);

describe('Production Manifest Validation', () => {
  let manifestContent: string;
  let manifestXml: any;
  const GITHUB_PAGES_URL = 'https://mikedtcm22.github.io/excel_conditional_fill';

  beforeAll(async () => {
    const manifestPath = path.join(__dirname, '../manifest-production.xml');
    manifestContent = fs.readFileSync(manifestPath, 'utf-8');
    manifestXml = await parseXml(manifestContent);
  });

  test('manifest-production.xml is valid XML and parseable', async () => {
    expect(manifestXml).toBeDefined();
    expect(manifestXml.OfficeApp).toBeDefined();
  });

  test('all URLs in production manifest point to GitHub Pages', () => {
    // Check IconUrl
    const iconUrl = manifestXml.OfficeApp.IconUrl[0].$.DefaultValue;
    expect(iconUrl).toContain(GITHUB_PAGES_URL);
    expect(iconUrl).not.toContain('localhost');

    // Check HighResolutionIconUrl
    const hiResIconUrl = manifestXml.OfficeApp.HighResolutionIconUrl[0].$.DefaultValue;
    expect(hiResIconUrl).toContain(GITHUB_PAGES_URL);

    // Check SourceLocation
    const sourceLocation =
      manifestXml.OfficeApp.DefaultSettings[0].SourceLocation[0].$.DefaultValue;
    expect(sourceLocation).toContain(GITHUB_PAGES_URL);
    expect(sourceLocation).toContain('taskpane.html');
  });

  test('context menu entries are properly defined', () => {
    const versionOverrides = manifestXml.OfficeApp.VersionOverrides[0];
    const desktopFormFactor = versionOverrides.Hosts[0].Host[0].DesktopFormFactor[0];
    const extensionPoints = desktopFormFactor.ExtensionPoint;

    // Find the ContextMenu extension point
    const contextMenuExtension = extensionPoints.find(
      (ep: any) => ep.$['xsi:type'] === 'ContextMenu'
    );
    expect(contextMenuExtension).toBeDefined();

    // Check that OfficeMenu with ContextMenuCell exists
    const officeMenu = contextMenuExtension.OfficeMenu[0];
    expect(officeMenu.$.id).toBe('ContextMenuCell');

    // Check that both Fill Blanks and Fill Errors controls exist
    const controls = officeMenu.Control;
    expect(controls.length).toBeGreaterThanOrEqual(2);

    const fillBlanksControl = controls.find(
      (c: any) => c.$.id === 'ContextMenuFillBlanks'
    );
    const fillErrorsControl = controls.find(
      (c: any) => c.$.id === 'ContextMenuFillErrors'
    );

    expect(fillBlanksControl).toBeDefined();
    expect(fillErrorsControl).toBeDefined();

    // Check action function names
    expect(fillBlanksControl.Action[0].FunctionName[0]).toBe('fillBlanksCommand');
    expect(fillErrorsControl.Action[0].FunctionName[0]).toBe('fillErrorsCommand');
  });

  test('ExtendedOverrides references shortcuts.json on GitHub Pages', () => {
    const versionOverrides = manifestXml.OfficeApp.VersionOverrides[0];
    const extendedOverrides = versionOverrides.ExtendedOverrides;

    expect(extendedOverrides).toBeDefined();
    expect(extendedOverrides[0].$.Url).toContain(GITHUB_PAGES_URL);
    expect(extendedOverrides[0].$.Url).toContain('shortcuts.json');
  });
});

describe('Development Manifest Context Menu', () => {
  let manifestXml: any;

  beforeAll(async () => {
    const manifestPath = path.join(__dirname, '../manifest.xml');
    const manifestContent = fs.readFileSync(manifestPath, 'utf-8');
    manifestXml = await parseXml(manifestContent);
  });

  test('development manifest also has context menu extension point', () => {
    const versionOverrides = manifestXml.OfficeApp.VersionOverrides[0];
    const desktopFormFactor = versionOverrides.Hosts[0].Host[0].DesktopFormFactor[0];
    const extensionPoints = desktopFormFactor.ExtensionPoint;

    const contextMenuExtension = extensionPoints.find(
      (ep: any) => ep.$['xsi:type'] === 'ContextMenu'
    );
    expect(contextMenuExtension).toBeDefined();
  });

  test('development manifest has ExtendedOverrides for localhost', () => {
    const versionOverrides = manifestXml.OfficeApp.VersionOverrides[0];
    const extendedOverrides = versionOverrides.ExtendedOverrides;

    expect(extendedOverrides).toBeDefined();
    expect(extendedOverrides[0].$.Url).toContain('localhost');
    expect(extendedOverrides[0].$.Url).toContain('shortcuts.json');
  });
});
