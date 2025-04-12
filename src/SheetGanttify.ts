import { stringify } from 'csv-stringify/sync';
import dayjs from 'dayjs';
import { TaskDefinition, resolveSchedule } from './task';
import { WorkdayCalendar } from './workday-calendar';

const GSS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_EDIT = GSS.getSheetByName('edit');
const SHEET_VIEW = GSS.getSheetByName('view');
const COL_CHART = 21;
const ROW_DATA = 5;

const MONTH = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12'];
const WEEK = ['日', '月', '火', '水', '木', '金', '土'];
const HOLIDAY_NATIONAL = '祝';
const HOLIDAY_USER = 'ユ';

export class SheetGanttify {
  // eslint-disable-next-line no-use-before-define
  private static instance: SheetGanttify;

  private _params:
    | {
        calendarStart: dayjs.Dayjs;
        calendarEnd: dayjs.Dayjs;
        calendarDuration: number;
        ticketTracker: string;
        ticketTargetVersion: string;
        ticketLinkBasePath: string;
      }
    | undefined;

  private _holidays:
    | {
        national: Set<string>;
        user: Set<string>;
      }
    | undefined;

  private _calendar: WorkdayCalendar | undefined;

  private _taskData:
    | {
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
      }
    | undefined;

  private constructor() {}

  public static getInstance(): SheetGanttify {
    if (!SheetGanttify.instance) {
      SheetGanttify.instance = new SheetGanttify();
    }
    return SheetGanttify.instance;
  }

  private get params() {
    if (this._params) return this._params;

    const params = Sheets.Spreadsheets?.Values?.batchGet(
      SpreadsheetApp.getActive().getId(),
      { ranges: ['params!2:2'] }
    )?.valueRanges?.[0]?.values;

    if (typeof params?.[0]?.[0] !== 'string')
      throw new Error('calendarStart is not defined');
    const calendarStart = dayjs(params?.[0]?.[0]);

    if (typeof params?.[0]?.[1] !== 'string')
      throw new Error('calendarEnd is not defined');
    const calendarEnd = dayjs(params?.[0]?.[1]);

    if (calendarStart > calendarEnd)
      throw new Error('calendarStart > calendarEnd');
    const calendarDuration = calendarEnd.diff(calendarStart, 'day') + 1;

    const ticketTracker =
      typeof params?.[0]?.[2] === 'string' ? params?.[0]?.[2] : '';

    const ticketTargetVersion =
      typeof params?.[0]?.[3] === 'string' ? params?.[0]?.[3] : '';

    const ticketLinkBasePath =
      typeof params?.[0]?.[4] === 'string' ? params?.[0]?.[4] : '';

    return (this._params = {
      calendarStart,
      calendarEnd,
      calendarDuration,
      ticketTracker,
      ticketTargetVersion,
      ticketLinkBasePath,
    });
  }

  private get holidays() {
    if (this._holidays) return this._holidays;

    const holidays = (
      Sheets.Spreadsheets?.Values?.batchGet(
        SpreadsheetApp.getActive().getId(),
        { ranges: ['holiday!A3:A', 'holiday!C3:C'], majorDimension: 'COLUMNS' }
      )?.valueRanges ?? []
    ).map(r =>
      (r.values ?? [[]])[0]
        .filter(v => typeof v === 'string')
        .filter(v => v !== '')
        .map(v => dayjs(v).format('YYYY/MM/DD'))
    );

    const national = new Set(holidays[0]);
    const user = new Set(holidays[1]);

    return (this._holidays = {
      national,
      user,
    });
  }

  private get calendar() {
    if (this._calendar) return this._calendar;

    const holidays = new Set([
      ...this.holidays.national,
      ...this.holidays.user,
    ]);
    return (this._calendar = new WorkdayCalendar(
      this.params.calendarStart.format('YYYY/MM/DD'),
      this.params.calendarDuration,
      (date, formatted) => {
        if (date.day() === 0 || date.day() === 6) return false;
        if (holidays.has(formatted)) return false;
        return true;
      }
    ));
  }

  private get taskData() {
    if (this._taskData) return this._taskData;

    this._taskData = {
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

    const data = Sheets.Spreadsheets?.Values?.batchGet(
      SpreadsheetApp.getActive().getId(),
      {
        ranges: [
          this._taskData.ticketId.range,
          this._taskData.sectionAndTask.range,
          this._taskData.start.range,
          this._taskData.end.range,
          this._taskData.actual.range,
          this._taskData.link.range,
          this._taskData.assignee.range,
          this._taskData.progress.range,
          this._taskData.state.range,
        ],
      }
    )?.valueRanges;
    if (data === undefined) throw new Error('Failed to get data');
    const formula = Sheets.Spreadsheets?.Values?.batchGet(
      SpreadsheetApp.getActive().getId(),
      {
        ranges: [this._taskData.start.range, this._taskData.end.range],
        valueRenderOption: 'FORMULA',
      }
    )?.valueRanges;
    if (formula === undefined) throw new Error('Failed to get formula');

    this._taskData.ticketId.data = (data[0].values ?? [[]]) as string[][];
    this._taskData.sectionAndTask.data = (data[1].values ?? [[]]) as string[][];
    this._taskData.start.data = (data[2].values ?? [[]]) as string[][];
    this._taskData.end.data = (data[3].values ?? [[]]) as string[][];
    this._taskData.actual.data = (data[4].values ?? [[]]) as string[][];
    this._taskData.link.data = (data[5].values ?? [[]]) as string[][];
    this._taskData.assignee.data = (data[6].values ?? [[]]) as string[][];
    this._taskData.progress.data = (data[7].values ?? [[]]) as string[][];
    this._taskData.state.data = (data[8].values ?? [[]]) as string[][];

    const startFormula = (formula[0].values ?? [[]]) as (string | number)[][];
    const endFormula = (formula[1].values ?? [[]]) as (string | number)[][];

    startFormula.forEach((row, index) => {
      if (
        typeof row[0] === 'string' &&
        row[0].startsWith('=') &&
        this._taskData?.start.data[index] !== undefined
      ) {
        this._taskData.start.data[index] = row
          .filter(col => typeof col === 'string')
          .filter(col => col.startsWith('='));
      }
    });
    endFormula.forEach((row, index) => {
      if (
        typeof row[0] === 'string' &&
        row[0].startsWith('=') &&
        this._taskData?.end.data[index] !== undefined
      ) {
        this._taskData.end.data[index] = row
          .filter(col => typeof col === 'string')
          .filter(col => col.startsWith('='));
      }
    });

    return this._taskData;
  }

  public parseDummy() {
    const taskData = this.taskData;

    const actual = Array.from(
      new Array(taskData.sectionAndTask.data.length),
      _ => new Array(2).fill('') as string[]
    );
    taskData.actual.data.forEach(([actualStart, actualEnd], index) => {
      if (actualStart) actual[index][0] = actualStart;
      if (actualEnd) actual[index][1] = actualEnd;
    });
    const assignee = Array.from(
      new Array(taskData.sectionAndTask.data.length),
      _ => new Array(1).fill('') as string[]
    );
    taskData.assignee.data.forEach(([a], index) => {
      if (a) assignee[index][0] = a;
    });
    const progress = Array.from(
      new Array(taskData.sectionAndTask.data.length),
      _ => new Array(1).fill('') as string[]
    );
    taskData.progress.data.forEach(([p], index) => {
      if (p) progress[index][0] = p;
    });

    taskData.ticketId.data.forEach(([t], index) => {
      if (!t || !t.startsWith('!')) return;

      const [p, s, e, a] = t.slice(1).split(',');
      if (p) progress[index][0] = p;
      if (s) actual[index][0] = s;
      if (e) actual[index][1] = e;
      if (a) assignee[index][0] = a;
    });

    taskData.actual.data = actual;
    taskData.assignee.data = assignee;
    taskData.progress.data = progress;

    taskData.actual.updated = true;
    taskData.assignee.updated = true;
    taskData.progress.updated = true;
  }

  public createLink() {
    const taskData = this.taskData;
    const params = this.params;

    taskData.link.data = taskData.sectionAndTask.data.map(
      ([_, task], index) => {
        if (!task) return [''];
        const ticketId = taskData.ticketId.data[index]?.[0];
        if (!ticketId || ticketId.startsWith('!')) return [task];
        return [
          `=HYPERLINK("${params.ticketLinkBasePath}${ticketId}", "${task.replace(/"/g, '""')}")`,
        ];
      }
    );
    taskData.link.updated = true;
  }

  public createGantt() {
    const taskData = this.taskData;

    const tasks: TaskDefinition[] = [];
    taskData.sectionAndTask.data.forEach(([_, task], index) => {
      if (task === undefined) return;

      let startDate: TaskDefinition['startDate'] = null;
      let endDate: TaskDefinition['endDate'] = null;
      let duration: TaskDefinition['duration'] = null;
      let startsAfter: TaskDefinition['startsAfter'] = new Set();
      let endsBefore: TaskDefinition['endsBefore'] = new Set();

      const s = taskData.start.data[index]?.[0];
      if (s === undefined) {
        // do nothing
      } else if (s.startsWith('=')) {
        startsAfter = new Set(
          taskData.start.data[index].map(
            v => parseInt(v.replace(/=[A-Z]*/, '')) - ROW_DATA
          )
        );
      } else if (s.match(/^[0-9]+d$/)) {
        duration = parseInt(s.replace(/d$/, ''));
      } else if (s.match(/^[0-9]{4}\/[0-9]{2}\/[0-9]{2}$/)) {
        startDate = s;
      }

      const e = taskData.end.data[index]?.[0];
      if (e === undefined) {
        // do nothing
      } else if (e.startsWith('=')) {
        endsBefore = new Set(
          taskData.end.data[index].map(
            v => parseInt(v.replace(/=[A-Z]*/, '')) - ROW_DATA
          )
        );
      } else if (e.match(/^[0-9]+d$/)) {
        duration = parseInt(e.replace(/d$/, ''));
      } else if (e.match(/^[0-9]{4}\/[0-9]{2}\/[0-9]{2}$/)) {
        endDate = e;
      }

      tasks.push({
        id: index,
        startDate,
        endDate,
        duration,
        startsAfter,
        endsBefore,
      });
    });

    const calendar = this.calendar;
    const [, scheduledTasks] = resolveSchedule(
      tasks.filter(t => t.startDate || t.endDate || t.duration),
      calendar
    );

    const actualTasks = tasks
      .filter(t => t.startDate || t.endDate || t.duration)
      .map(task => {
        const actualTask = { ...task };

        const actualStart = taskData.actual.data[task.id]?.[0] || null;
        const actualEnd = taskData.actual.data[task.id]?.[1] || null;

        if (actualStart === null && actualEnd === null) {
          return actualTask;
        }
        if (actualStart && actualEnd) {
          actualTask.startDate = actualStart;
          actualTask.endDate = actualEnd;
          actualTask.duration = null;
          actualTask.startsAfter = new Set();
          actualTask.endsBefore = new Set();
          return actualTask;
        }
        if (actualTask.startDate && actualTask.endDate) {
          if (actualStart && actualStart <= actualTask.endDate) {
            actualTask.startDate = actualStart;
          }
          if (actualEnd && actualTask.startDate <= actualEnd) {
            actualTask.endDate = actualEnd;
          }
        } else {
          actualTask.startDate = actualStart;
          actualTask.endDate = actualEnd;
          actualTask.startsAfter = new Set();
          actualTask.endsBefore = new Set();
        }
        return actualTask;
      });
    const [, actualScheduledTasks] = resolveSchedule(actualTasks, calendar);

    const params = this.params;
    taskData.gantt.data = Array.from(
      new Array(tasks[tasks.length - 1].id + 1),
      _ => new Array(params.calendarDuration).fill('') as string[]
    );
    taskData.state.data = Array.from(
      new Array(tasks[tasks.length - 1].id + 1),
      _ => new Array(1).fill('') as string[]
    );

    const today = dayjs().format('YYYY/MM/DD');
    tasks.forEach(task => {
      const scheduledTask = scheduledTasks.get(task.id);
      const actualScheduledTask = actualScheduledTasks.get(task.id);
      let [actualStart, actualEnd] = taskData.actual.data[task.id] ?? [];
      const progress = taskData.progress.data[task.id]?.[0] ?? '';

      taskData.state.data[task.id][0] = '';
      if (actualStart) taskData.state.data[task.id][0] = '進行中';

      if (scheduledTask) {
        taskData.gantt.data[task.id].fill(
          '━',
          dayjs(scheduledTask.startDate).diff(params.calendarStart, 'day'),
          dayjs(scheduledTask.endDate).diff(params.calendarStart, 'day') + 1
        );
      }

      if (actualScheduledTask) {
        actualStart = actualScheduledTask.startDate!;
        actualEnd = actualScheduledTask.endDate!;
      }
      if (actualStart && actualEnd) {
        if (scheduledTask) {
          const scheduledStartIndex = dayjs(scheduledTask.startDate).diff(
            params.calendarStart,
            'day'
          );
          const scheduledEndIndex = dayjs(scheduledTask.endDate).diff(
            params.calendarStart,
            'day'
          );

          for (
            let i = dayjs(actualStart).diff(params.calendarStart, 'day');
            i < dayjs(actualEnd).diff(params.calendarStart, 'day') + 1;
            i++
          ) {
            if (i < scheduledStartIndex) {
              taskData.gantt.data[task.id][i] = '┫';
            } else if (i > scheduledEndIndex) {
              taskData.gantt.data[task.id][i] = '┣';
              taskData.state.data[task.id][0] = '遅延';
            } else {
              taskData.gantt.data[task.id][i] = '╋';
            }
          }
        } else {
          taskData.gantt.data[task.id].fill(
            '┃',
            dayjs(actualStart).diff(params.calendarStart, 'day'),
            dayjs(actualEnd).diff(params.calendarStart, 'day') + 1
          );
        }
      }
      if (actualScheduledTask && !taskData.actual.data[task.id]?.[0]) {
        if (today >= actualScheduledTask.startDate!)
          taskData.state.data[task.id][0] = '要開始';
        else taskData.state.data[task.id][0] = '未着手';
      }
      if (
        taskData.actual.data[task.id]?.[0] &&
        progress !== '100' &&
        actualEnd < today
      ) {
        for (
          let i = dayjs(actualEnd).diff(params.calendarStart, 'day') + 1;
          i < dayjs(today).diff(params.calendarStart, 'day') + 1;
          i++
        ) {
          taskData.gantt.data[task.id][i] = '╳';
        }
        taskData.state.data[task.id][0] = '要見直';
      }
      if (progress === '100') taskData.state.data[task.id][0] = '完了';
    });

    taskData.gantt.updated = true;
    taskData.state.updated = true;
  }

  public write() {
    const taskData = this.taskData;

    const data = Object.values(taskData)
      .filter(d => d.updated)
      .map(({ range, data }) => ({
        range,
        values: data,
      }));
    if (data.length === 0) return;

    Sheets.Spreadsheets?.Values?.batchUpdate(
      {
        valueInputOption: 'USER_ENTERED',
        data,
      },
      SpreadsheetApp.getActive().getId()
    );
  }

  public createCalendar() {
    const params = this.params;
    const holidays = this.holidays;

    if (SHEET_VIEW) {
      Sheets.Spreadsheets?.Values?.batchClear(
        { ranges: ['view'] },
        SpreadsheetApp.getActive().getId()
      );
    }

    if (SHEET_EDIT && SHEET_EDIT.getMaxColumns() > COL_CHART) {
      SHEET_EDIT.deleteColumns(
        COL_CHART + 1,
        SHEET_EDIT.getMaxColumns() - COL_CHART
      );
    }
    if (SHEET_VIEW && SHEET_VIEW.getMaxColumns() > COL_CHART) {
      SHEET_VIEW.deleteColumns(
        COL_CHART + 1,
        SHEET_VIEW.getMaxColumns() - COL_CHART
      );
    }
    SHEET_EDIT?.insertColumnsAfter(COL_CHART, params.calendarDuration - 1);
    SHEET_VIEW?.insertColumnsAfter(COL_CHART, params.calendarDuration - 1);

    const months = [];
    const weeks = [];
    const days = [];
    const notes = [];
    for (let i = 0; i < params.calendarDuration; i++) {
      const day = params.calendarStart.add(i, 'day');
      months.push(MONTH[day.month()]);
      weeks.push(WEEK[day.day()]);
      days.push(day.format('DD'));
      const formatted = day.format('YYYY/MM/DD');
      notes.push(
        holidays.user.has(formatted)
          ? HOLIDAY_USER
          : holidays.national.has(formatted)
            ? HOLIDAY_NATIONAL
            : ''
      );
    }
    if (SHEET_EDIT) {
      Sheets.Spreadsheets?.Values?.batchUpdate(
        {
          valueInputOption: 'USER_ENTERED',
          data: [{ range: 'edit!U1:4', values: [months, weeks, days, notes] }],
        },
        SpreadsheetApp.getActive().getId()
      );
    }

    months.push('end');
    let mergeStartColumn = COL_CHART;
    let mergeColNum = 1;
    for (let i = 0; i < months.length - 1; i++) {
      const month = months[i];
      const nextMonth = months[i + 1];
      if (month === nextMonth) {
        mergeColNum++;
      } else {
        SHEET_EDIT?.getRange(1, mergeStartColumn, 1, mergeColNum).merge();
        SHEET_VIEW?.getRange(1, mergeStartColumn, 1, mergeColNum).merge();
        mergeStartColumn = mergeStartColumn + mergeColNum;
        mergeColNum = 1;
      }
    }

    if (SHEET_EDIT && SHEET_VIEW) {
      Sheets.Spreadsheets?.Values?.batchUpdate(
        {
          valueInputOption: 'USER_ENTERED',
          data: [
            {
              range: 'view!A1:A1',
              values: [
                [
                  `=ARRAYFORMULA(edit!A1:${SHEET_EDIT.getRange(
                    1,
                    SHEET_EDIT.getMaxColumns()
                  )
                    .getA1Notation()
                    .replace(/\d/, '')})`,
                ],
              ],
            },
          ],
        },
        SpreadsheetApp.getActive().getId()
      );
    }
  }

  public generateCsvForImport() {
    const taskData = this.taskData;
    const params = this.params;

    const data: [string, string, string][] = [];

    taskData.sectionAndTask.data.forEach(([_, task], index) => {
      if (!task || taskData.ticketId.data[index]?.[0]) return;

      data.push([params.ticketTargetVersion, params.ticketTracker, task]);
    });

    return stringify([['対象バージョン', 'トラッカー', '題名'], ...data]);
  }

  public importTicketIds(ids: string[]) {
    this.parseDummy();
    const taskData = this.taskData;

    const ticketId = Array.from(
      new Array(taskData.sectionAndTask.data.length),
      _ => new Array(1).fill('') as string[]
    );
    taskData.ticketId.data.forEach(([t], index) => {
      if (t) ticketId[index][0] = t;
    });

    let i = 0;
    taskData.sectionAndTask.data.forEach(([_, task], index) => {
      if (!task) return;

      if (!ticketId[index][0]) {
        if (i < ids.length) ticketId[index][0] = ids[i++];
        else throw new Error('invalid input');
      }
    });
    if (i !== ids.length) throw new Error('invalid input');

    this.taskData.ticketId.data = ticketId;

    this.taskData.ticketId.updated = true;
  }

  public importInfo(info: Record<string, [string, string, string, string]>) {
    this.parseDummy();
    const taskData = this.taskData;

    taskData.sectionAndTask.data.forEach(([_, task], index) => {
      if (!task) return;

      const ticketId = taskData.ticketId.data[index][0];

      if (ticketId !== '' && !ticketId.startsWith('!') && ticketId in info) {
        taskData.progress.data[index][0] = info[ticketId][0];
        taskData.actual.data[index][0] = info[ticketId][1];
        taskData.actual.data[index][1] = info[ticketId][2];
        taskData.assignee.data[index][0] = info[ticketId][3];
      }
    });
  }
}
