/* eslint-disable @typescript-eslint/no-unsafe-assignment, @typescript-eslint/no-explicit-any, @typescript-eslint/no-unsafe-member-access, @typescript-eslint/no-require-imports */
import { ROW_DATA } from '../src/constants';
import { TaskData } from '../src/sheet-task-loader';
import { TaskDefinition } from '../src/task';

let SheetGanttify: typeof import('../src/SheetGanttify').SheetGanttify;
let instance: {
  _parseDuration: (v: unknown) => number | null;
  _parseDate: (v: unknown) => string | null;
  _parseTaskDefinitions: () => TaskDefinition[];
  _taskData: TaskData;
};
beforeAll(() => {
  (global as any).SpreadsheetApp = {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => null,
      getId: () => 'dummy',
    }),
    getActive: () => ({ getId: () => 'dummy' }),
  };
  (global as any).Sheets = {
    Spreadsheets: {
      Values: {
        batchGet: jest.fn(() => ({})),
        batchUpdate: jest.fn(() => ({})),
        batchClear: jest.fn(() => ({})),
      },
    },
  };
  SheetGanttify = require('../src/SheetGanttify').SheetGanttify;
  instance = SheetGanttify.getInstance() as unknown as {
    _parseDuration: (v: unknown) => number | null;
    _parseDate: (v: unknown) => string | null;
    _parseTaskDefinitions: () => TaskDefinition[];
    _taskData: TaskData;
  };
});

describe('SheetGanttify parser helpers', () => {
  test('_parseDuration', () => {
    expect(instance._parseDuration('3d')).toBe(3);
    expect(instance._parseDuration('10d')).toBe(10);
    expect(instance._parseDuration('3')).toBeNull();
    expect(instance._parseDuration(5)).toBeNull();
  });

  test('_parseDate', () => {
    expect(instance._parseDate('2025/01/01')).toBe('2025/01/01');
    expect(instance._parseDate('2025-01-01')).toBeNull();
    expect(instance._parseDate('invalid')).toBeNull();
  });

  test('_parseTaskDefinitions', () => {
    const dummy: TaskData = {
      ticketId: { range: '', data: [[''], [''], ['']], updated: false },
      sectionAndTask: {
        range: '',
        data: [
          ['', 'Task1'],
          ['', ''],
          ['', 'Task2'],
        ],
        updated: false,
      },
      start: {
        range: '',
        data: [['2025/01/01'], [''], ['=T' + (ROW_DATA + 0)]],
        updated: false,
      },
      end: { range: '', data: [[''], [''], ['']], updated: false },
      actual: { range: '', data: [[''], [''], ['']], updated: false },
      link: { range: '', data: [[''], [''], ['']], updated: false },
      assignee: { range: '', data: [[''], [''], ['']], updated: false },
      progress: { range: '', data: [[''], [''], ['']], updated: false },
      state: { range: '', data: [[''], [''], ['']], updated: false },
      gantt: { range: '', data: [[''], [''], ['']], updated: false },
    };
    instance._taskData = dummy;

    const result = instance._parseTaskDefinitions();
    expect(result).toEqual([
      {
        id: 0,
        startDate: '2025/01/01',
        endDate: null,
        duration: null,
        startsAfter: new Set(),
        endsBefore: new Set(),
      },
      {
        id: 2,
        startDate: null,
        endDate: null,
        duration: null,
        startsAfter: new Set([0]),
        endsBefore: new Set(),
      },
    ]);
  });
});
