import * as fs from 'fs';
import * as path from 'path';

describe('Task Pane TypeScript Initialization', () => {
  let taskpaneCode: string;

  beforeAll(() => {
    const taskpanePath = path.join(__dirname, '../src/taskpane/taskpane.ts');
    taskpaneCode = fs.readFileSync(taskpanePath, 'utf-8');
  });

  test('taskpane.ts contains Office.onReady or Office.initialize', () => {
    const hasOfficeReady = taskpaneCode.includes('Office.onReady') || taskpaneCode.includes('Office.initialize');
    expect(hasOfficeReady).toBe(true);
  });

  test('taskpane.ts contains getTargetCondition function', () => {
    expect(taskpaneCode).toContain('getTargetCondition');
  });

  test('taskpane.ts contains getTemplateSource function', () => {
    expect(taskpaneCode).toContain('getTemplateSource');
  });

  test('taskpane.ts contains Run button event handler setup', () => {
    const hasRunButtonHandler = taskpaneCode.includes('runButton') &&
                                (taskpaneCode.includes('addEventListener') || taskpaneCode.includes('onclick'));
    expect(hasRunButtonHandler).toBe(true);
  });

  test('taskpane.ts contains updateStatus function', () => {
    expect(taskpaneCode).toContain('updateStatus');
  });

  test('taskpane.ts contains runFillOperation placeholder or call', () => {
    expect(taskpaneCode).toContain('runFillOperation');
  });
});
