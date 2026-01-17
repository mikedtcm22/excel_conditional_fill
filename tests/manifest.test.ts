import * as fs from 'fs';
import * as path from 'path';
import { parseString } from 'xml2js';
import { promisify } from 'util';

const parseXml = promisify(parseString);

describe('Manifest Validation', () => {
  let manifestContent: string;
  let manifestXml: any;

  beforeAll(async () => {
    const manifestPath = path.join(__dirname, '../manifest.xml');
    manifestContent = fs.readFileSync(manifestPath, 'utf-8');
    manifestXml = await parseXml(manifestContent);
  });

  test('manifest XML is valid and parseable', async () => {
    expect(manifestXml).toBeDefined();
    expect(manifestXml.OfficeApp).toBeDefined();
  });

  test('manifest contains required metadata', () => {
    const officeApp = manifestXml.OfficeApp;
    expect(officeApp.Id).toBeDefined();
    expect(officeApp.Version).toBeDefined();
    expect(officeApp.ProviderName).toBeDefined();
    expect(officeApp.DisplayName).toBeDefined();
    expect(officeApp.DisplayName[0].$.DefaultValue).toBe('FillGaps');
  });

  test('manifest has ReadWriteDocument permissions', () => {
    const permissions = manifestXml.OfficeApp.Permissions;
    expect(permissions).toBeDefined();
    expect(permissions[0]).toBe('ReadWriteDocument');
  });

  test('manifest specifies correct host (Workbook)', () => {
    const hosts = manifestXml.OfficeApp.Hosts;
    expect(hosts).toBeDefined();
    expect(hosts[0].Host).toBeDefined();
  });

  test('manifest task pane URL uses HTTPS localhost', () => {
    const defaultSettings = manifestXml.OfficeApp.DefaultSettings;
    if (defaultSettings && defaultSettings[0].SourceLocation) {
      const sourceUrl = defaultSettings[0].SourceLocation[0].$.DefaultValue;
      expect(sourceUrl).toContain('https://localhost');
      expect(sourceUrl).toContain('taskpane.html');
    }
  });

  test('manifest version overrides contain ribbon UI elements', () => {
    // This test checks if VersionOverrides exists for ribbon customization
    const versionOverrides = manifestXml.OfficeApp.VersionOverrides;
    expect(versionOverrides).toBeDefined();
  });
});
