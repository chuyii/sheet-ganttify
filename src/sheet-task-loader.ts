import { ROW_DATA } from './constants';

const GSS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_EDIT = GSS.getSheetByName('edit');

export type TaskData = {
  ticketId: { range: string; data: string[][]; updated: boolean };
  sectionAndTask: { range: string; data: string[][]; updated: boolean };
  start: { range: string; data: string[][]; updated: boolean };
  end: { range: string; data: string[][]; updated: boolean };
  actual: { range: string; data: string[][]; updated: boolean };
  link: { range: string; data: string[][]; updated: boolean };
  assignee: { range: string; data: string[][]; updated: boolean };
  progress: { range: string; data: string[][]; updated: boolean };
  state: { range: string; data: string[][]; updated: boolean };
  gantt: { range: string; data: string[][]; updated: boolean };
};

export class SheetTaskLoader {
  public static load(): TaskData {
    const taskData: TaskData = {
      ticketId: { range: `edit!B${ROW_DATA}:B`, data: [[]], updated: false },
      sectionAndTask: {
        range: `edit!C${ROW_DATA}:D`,
        data: [[]],
        updated: false,
      },
      start: { range: `edit!E${ROW_DATA}:I`, data: [[]], updated: false },
      end: { range: `edit!J${ROW_DATA}:N`, data: [[]], updated: false },
      actual: { range: `edit!O${ROW_DATA}:P`, data: [[]], updated: false },
      link: { range: `edit!Q${ROW_DATA}:Q`, data: [[]], updated: false },
      assignee: { range: `edit!R${ROW_DATA}:R`, data: [[]], updated: false },
      progress: { range: `edit!S${ROW_DATA}:S`, data: [[]], updated: false },
      state: { range: `edit!T${ROW_DATA}:T`, data: [[]], updated: false },
      gantt: {
        range: `edit!U${ROW_DATA}:${SHEET_EDIT!
          .getRange(1, SHEET_EDIT!.getMaxColumns())
          .getA1Notation()
          .replace(/\d/, '')}T`,
        data: [[]],
        updated: false,
      },
    };

    const ranges = [
      taskData.ticketId.range,
      taskData.sectionAndTask.range,
      taskData.start.range,
      taskData.end.range,
      taskData.actual.range,
      taskData.link.range,
      taskData.assignee.range,
      taskData.progress.range,
      taskData.state.range,
    ];

    const data = Sheets.Spreadsheets?.Values?.batchGet(
      SpreadsheetApp.getActive().getId(),
      { ranges }
    )?.valueRanges;
    if (data === undefined) throw new Error('Failed to get data');

    const formula = Sheets.Spreadsheets?.Values?.batchGet(
      SpreadsheetApp.getActive().getId(),
      {
        ranges: [taskData.start.range, taskData.end.range],
        valueRenderOption: 'FORMULA',
      }
    )?.valueRanges;
    if (formula === undefined) throw new Error('Failed to get formula');

    taskData.ticketId.data = (data[0].values ?? [[]]) as string[][];
    taskData.sectionAndTask.data = (data[1].values ?? [[]]) as string[][];
    taskData.start.data = (data[2].values ?? [[]]) as string[][];
    taskData.end.data = (data[3].values ?? [[]]) as string[][];
    taskData.actual.data = (data[4].values ?? [[]]) as string[][];
    taskData.link.data = (data[5].values ?? [[]]) as string[][];
    taskData.assignee.data = (data[6].values ?? [[]]) as string[][];
    taskData.progress.data = (data[7].values ?? [[]]) as string[][];
    taskData.state.data = (data[8].values ?? [[]]) as string[][];

    const startFormula = (formula[0].values ?? [[]]) as (string | number)[][];
    const endFormula = (formula[1].values ?? [[]]) as (string | number)[][];

    startFormula.forEach((row, index) => {
      if (
        typeof row[0] === 'string' &&
        row[0].startsWith('=') &&
        taskData.start.data[index] !== undefined
      ) {
        taskData.start.data[index] = row
          .filter(col => typeof col === 'string')
          .filter(col => col.startsWith('='));
      }
    });
    endFormula.forEach((row, index) => {
      if (
        typeof row[0] === 'string' &&
        row[0].startsWith('=') &&
        taskData.end.data[index] !== undefined
      ) {
        taskData.end.data[index] = row
          .filter(col => typeof col === 'string')
          .filter(col => col.startsWith('='));
      }
    });

    return taskData;
  }
}
