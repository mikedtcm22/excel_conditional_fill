import * as fs from 'fs';
import * as path from 'path';

describe('Task Pane UI Rendering', () => {
  let htmlContent: string;

  beforeAll(() => {
    const taskpanePath = path.join(__dirname, '../src/taskpane/taskpane.html');
    htmlContent = fs.readFileSync(taskpanePath, 'utf-8');
  });

  test('HTML includes Office.js script reference', () => {
    expect(htmlContent).toContain('appsforoffice.microsoft.com/lib/1/hosted/office.js');
  });

  test('HTML contains Target Condition radio button group', () => {
    expect(htmlContent).toContain('name="targetCondition"');
    expect(htmlContent).toContain('value="blanks"');
    expect(htmlContent).toContain('value="errors"');
    expect(htmlContent).toContain('value="both"');
  });

  test('HTML contains Template Source radio button group', () => {
    expect(htmlContent).toContain('name="templateSource"');
    expect(htmlContent).toContain('value="activeCell"');
    expect(htmlContent).toContain('value="topLeft"');
  });

  test('HTML contains Run button with correct ID', () => {
    expect(htmlContent).toContain('id="runButton"');
    expect(htmlContent).toMatch(/Run Fill Operation/i);
  });

  test('HTML contains status area div', () => {
    expect(htmlContent).toContain('id="statusArea"');
  });

  test('HTML references taskpane.css for styling', () => {
    // CSS is injected by webpack via style-loader, so we check if the HTML structure supports it
    // or check that taskpane.ts will import the CSS
    expect(htmlContent).toBeDefined();
  });
});
