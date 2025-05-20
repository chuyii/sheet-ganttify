import { stringify } from 'csv-stringify/sync';
import dayjs from 'dayjs';
import {
  COL_CHART,
  HOLIDAY_NATIONAL,
  HOLIDAY_USER,
  MONTH,
  ROW_DATA,
  STATUS,
  WEEK,
} from './constants';
import { SheetTaskLoader, TaskData } from './sheet-task-loader';
import { TaskDefinition, resolveSchedule } from './task';
import { WorkdayCalendar } from './workday-calendar';

const GSS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_EDIT = GSS.getSheetByName('edit');
const SHEET_VIEW = GSS.getSheetByName('view');

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

  private _taskData: TaskData | undefined;

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

    this._taskData = SheetTaskLoader.load();

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

  private _parseDependencies(info: (string | number)[]): Set<number> {
    return new Set(
      info
        .map(v => String(v).replace(/=[A-Z]*/, ''))
        .map(v => parseInt(v))
        .filter(v => !isNaN(v))
        .map(v => v - ROW_DATA)
        .filter(v => v >= 0)
    );
  }

  private _parseDuration(value: unknown): number | null {
    return typeof value === 'string' && /^[0-9]+d$/.test(value)
      ? parseInt(value.replace(/d$/, ''))
      : null;
  }

  private _parseDate(value: unknown): string | null {
    return typeof value === 'string' &&
      dayjs(value).isValid() &&
      /^\d{4}\/\d{2}\/\d{2}$/.test(value)
      ? value
      : null;
  }

  /**
   * シートのデータからタスク定義の配列を生成する
   * @returns TaskDefinition[]
   */
  private _parseTaskDefinitions(): TaskDefinition[] {
    const taskData = this.taskData;
    const tasks: TaskDefinition[] = [];

    const numRows = taskData.sectionAndTask.data.length;

    for (let index = 0; index < numRows; index++) {
      const taskRow = taskData.sectionAndTask.data[index];
      if (taskRow[1] === undefined || taskRow[1] === '') {
        continue;
      }

      let startDate: TaskDefinition['startDate'] = null;
      let endDate: TaskDefinition['endDate'] = null;
      let duration: TaskDefinition['duration'] = null;
      let startsAfter: TaskDefinition['startsAfter'] = new Set();
      let endsBefore: TaskDefinition['endsBefore'] = new Set();

      const startInfo = taskData.start.data[index] ?? [];
      const endInfo = taskData.end.data[index] ?? [];

      // --- Parse Start Info ---
      const s = startInfo[0];
      if (s !== undefined && s !== '') {
        if (typeof s === 'string' && s.startsWith('=')) {
          startsAfter = this._parseDependencies(startInfo);
        } else {
          const dur = this._parseDuration(s);
          const date = dur === null ? this._parseDate(s) : null;
          if (dur !== null) duration = dur;
          if (date) startDate = date;
        }
      }

      // --- Parse End Info ---
      const e = endInfo[0];
      if (e !== undefined && e !== '') {
        if (typeof e === 'string' && e.startsWith('=')) {
          endsBefore = this._parseDependencies(endInfo);
        } else {
          const dur = this._parseDuration(e);
          const date = dur === null ? this._parseDate(e) : null;
          if (dur !== null && duration === null) duration = dur;
          if (date) endDate = date;
        }
      }

      tasks.push({
        id: index,
        startDate,
        endDate,
        duration,
        startsAfter,
        endsBefore,
      });
    }
    return tasks;
  }

  /**
   * Prepare task definitions with actual data applied.
   * @param baseTasks TaskDefinition[] - parsed original task definitions
   * @returns TaskDefinition[] - task definitions updated with actual data
   */
  private _prepareActualTaskDefinitions(
    baseTasks: TaskDefinition[]
  ): TaskDefinition[] {
    const taskData = this.taskData;
    return baseTasks.map(task => {
      // Create a copy to avoid modifying the original base task definition
      const actualTask = {
        ...task,
        startsAfter: new Set(task.startsAfter), // Ensure sets are copied
        endsBefore: new Set(task.endsBefore),
      };

      const actualStartRaw = taskData.actual.data[task.id]?.[0];
      const actualEndRaw = taskData.actual.data[task.id]?.[1];

      // Validate actual dates before using them
      const actualStart =
        actualStartRaw &&
        typeof actualStartRaw === 'string' &&
        dayjs(actualStartRaw).isValid() &&
        actualStartRaw.match(/^\d{4}\/\d{2}\/\d{2}$/)
          ? actualStartRaw
          : null;
      const actualEnd =
        actualEndRaw &&
        typeof actualEndRaw === 'string' &&
        dayjs(actualEndRaw).isValid() &&
        actualEndRaw.match(/^\d{4}\/\d{2}\/\d{2}$/)
          ? actualEndRaw
          : null;

      // If neither actual start nor end date is valid, return the original task definition
      if (actualStart === null && actualEnd === null) {
        return actualTask;
      }

      // If both actual start and end dates are valid, override completely
      if (actualStart && actualEnd) {
        actualTask.startDate = actualStart;
        actualTask.endDate = actualEnd;
        actualTask.duration = null; // Duration is now implicitly defined by dates
        actualTask.startsAfter = new Set(); // Dependencies are overridden by actual dates
        actualTask.endsBefore = new Set();
        return actualTask;
      }

      // --- Handle cases where only one actual date is provided ---

      // If the original task was defined by fixed dates (start and end)
      if (actualTask.startDate && actualTask.endDate) {
        if (actualStart && actualStart <= actualTask.endDate) {
          // If actual start is provided and is on or before the original end date
          actualTask.startDate = actualStart; // Adjust start date
        }
        if (actualEnd && actualEnd >= actualTask.startDate) {
          // If actual end is provided and is on or after the (potentially updated) start date
          actualTask.endDate = actualEnd; // Adjust end date
        }
      } else {
        // If the original task was defined by duration or dependencies
        // We now have at least one fixed date (actualStart or actualEnd)
        actualTask.startDate = actualStart; // Set start if available
        actualTask.endDate = actualEnd; // Set end if available
        actualTask.startsAfter = new Set();
        actualTask.endsBefore = new Set();
      }

      return actualTask;
    });
  }

  /**
   * Compute planned and actual schedules.
   * @param tasks TaskDefinition[]
   * @returns [Map<number, ScheduledTask>, Map<number, ScheduledTask>] - [planned schedule, actual schedule]
   */
  private _resolveSchedules(
    tasks: TaskDefinition[]
  ): [Map<number, TaskDefinition>, Map<number, TaskDefinition>] {
    const calendar = this.calendar;
    // Filter tasks that have enough information to be scheduled
    const schedulableTasks = tasks.filter(
      t => t.startDate || t.endDate || t.duration
    );

    // Resolve planned schedule
    const [, scheduledTasks] = resolveSchedule(schedulableTasks, calendar);

    // Prepare task definitions reflecting actual start/end dates from the sheet
    const actualTaskDefinitions =
      this._prepareActualTaskDefinitions(schedulableTasks);

    // Resolve schedule based on actual task definitions
    const [, actualScheduledTasks] = resolveSchedule(
      actualTaskDefinitions,
      calendar
    );

    return [scheduledTasks, actualScheduledTasks];
  }

  /**
   * Generate gantt chart and status data.
   * @param tasks TaskDefinition[] - original task definitions used for mapping IDs and actual data
   * @param scheduledTasks Map<number, ScheduledTask> - planned schedule
   * @param actualScheduledTasks Map<number, ScheduledTask> - actual schedule
   */
  private _generateGanttAndStateData(
    tasks: TaskDefinition[], // Use the original parsed tasks list for consistent indexing
    scheduledTasks: Map<number, TaskDefinition>,
    actualScheduledTasks: Map<number, TaskDefinition>
  ) {
    const taskData = this.taskData;
    const params = this.params;

    // Determine the required array size based on the highest task ID + 1 from original parsing
    const maxId = tasks.length > 0 ? Math.max(...tasks.map(t => t.id)) : -1;
    const arraySize = maxId >= 0 ? maxId + 1 : 0; // Use 0 if no tasks

    // Initialize arrays based on maxId to ensure all task rows are covered
    const ganttData = Array.from(
      new Array(arraySize),
      () => new Array(params.calendarDuration).fill('') as string[]
    );
    const stateData = Array.from(
      new Array(arraySize),
      () => new Array(1).fill('') as string[]
    );

    const today = dayjs().format('YYYY/MM/DD');

    tasks.forEach(task => {
      const taskId = task.id;
      const scheduledTask = scheduledTasks.get(taskId);
      const actualScheduledTask = actualScheduledTasks.get(taskId);
      const [rawActualStart] = taskData.actual.data[taskId] ?? [];
      const progress = taskData.progress.data[taskId]?.[0] ?? '';

      // --- Initialize State ---
      let currentState = ''; // Default empty state
      if (rawActualStart) {
        // If there's any entry in Actual Start
        currentState = STATUS.IN_PROGRESS;
      }

      // --- Draw Planned Bar ---
      if (scheduledTask) {
        const startIndex = dayjs(scheduledTask.startDate).diff(
          params.calendarStart,
          'day'
        );
        const endIndex = dayjs(scheduledTask.endDate).diff(
          params.calendarStart,
          'day'
        );
        const validStartIndex = Math.max(0, startIndex);
        const validEndIndex = Math.min(params.calendarDuration - 1, endIndex);

        if (validStartIndex <= validEndIndex) {
          for (let i = validStartIndex; i <= validEndIndex; i++) {
            ganttData[taskId][i] = '━'; // Planned bar segment
          }
        }
      }

      // --- Draw Actual Bar (overriding planned bar segments) ---
      const actualStart = actualScheduledTask?.startDate ?? null;
      const actualEnd = actualScheduledTask?.endDate ?? null;

      if (actualStart && actualEnd) {
        const actualStartIndex = dayjs(actualStart).diff(
          params.calendarStart,
          'day'
        );
        const actualEndIndex = dayjs(actualEnd).diff(
          params.calendarStart,
          'day'
        );
        const validActualStartIndex = Math.max(0, actualStartIndex);
        const validActualEndIndex = Math.min(
          params.calendarDuration - 1,
          actualEndIndex
        );

        if (validActualStartIndex <= validActualEndIndex) {
          const scheduledStartIndex = scheduledTask
            ? dayjs(scheduledTask.startDate).diff(params.calendarStart, 'day')
            : -Infinity;
          const scheduledEndIndex = scheduledTask
            ? dayjs(scheduledTask.endDate).diff(params.calendarStart, 'day')
            : -Infinity;

          for (let i = validActualStartIndex; i <= validActualEndIndex; i++) {
            if (i < scheduledStartIndex) {
              ganttData[taskId][i] = '┫'; // Actual started before planned
            } else if (i > scheduledEndIndex) {
              ganttData[taskId][i] = '┣'; // Actual ended after planned
              if (currentState !== STATUS.DONE) currentState = STATUS.DELAYED; // Mark as delayed if not completed
            } else if (scheduledTask) {
              // Check scheduledTask exists before marking overlap
              ganttData[taskId][i] = '╋'; // Actual overlaps with planned
            } else {
              ganttData[taskId][i] = '┃'; // Actual bar, no planned bar to compare
            }
          }
          // If actual end is after planned end, ensure state reflects delay
          if (
            actualEndIndex > scheduledEndIndex &&
            currentState !== STATUS.DONE
          ) {
            currentState = STATUS.DELAYED;
          }
        }
      }

      // --- Update State based on Dates and Progress ---
      if (actualScheduledTask && !rawActualStart) {
        // Scheduled but not started according to raw data
        if (today >= actualScheduledTask.startDate!) {
          currentState = STATUS.NEED_START; // Past planned start, not started
        } else {
          currentState = STATUS.NOT_STARTED; // Not yet planned start
        }
      }

      // Check for overdue state: Started, not 100% done, and past *actual* end date.
      if (
        rawActualStart &&
        progress !== '100' &&
        actualEnd &&
        actualEnd < today
      ) {
        currentState = STATUS.NEED_REVIEW; // Task is overdue
        const actualEndIndex = dayjs(actualEnd).diff(
          params.calendarStart,
          'day'
        );
        const todayIndex = dayjs(today).diff(params.calendarStart, 'day');

        // Mark days from actualEnd+1 to today with '╳' if they are within the calendar range
        for (let i = actualEndIndex + 1; i <= todayIndex; i++) {
          if (i >= 0 && i < params.calendarDuration) {
            ganttData[taskId][i] = '╳';
          }
        }
      }

      // Final state override: If progress is 100%, it's '完了' regardless of dates.
      if (progress === '100') {
        currentState = STATUS.DONE;
      }

      // Assign the final calculated state
      stateData[taskId][0] = currentState;
    });

    // Assign the generated data back to taskData and mark as updated
    taskData.gantt.data = ganttData;
    taskData.state.data = stateData;
    taskData.gantt.updated = true;
    taskData.state.updated = true;
  }

  public createGantt() {
    // 1. Parse task definitions from the sheet data
    const taskDefinitions = this._parseTaskDefinitions();

    // If no valid tasks found, clear relevant data and exit
    if (taskDefinitions.length === 0) {
      const taskData = this.taskData;
      const params = this.params;
      // Ensure gantt/state arrays match the number of rows fetched initially
      const numRows = taskData.sectionAndTask.data.length;
      taskData.gantt.data = Array.from(
        new Array(numRows),
        _ => new Array(params.calendarDuration).fill('') as string[]
      );
      taskData.state.data = Array.from(
        new Array(numRows),
        _ => new Array(1).fill('') as string[]
      );
      taskData.gantt.updated = true;
      taskData.state.updated = true;
      return;
    }

    // 2. Resolve planned and actual schedules
    const [scheduledTasks, actualScheduledTasks] =
      this._resolveSchedules(taskDefinitions);

    // 3. Generate gantt chart visuals and task states based on schedules
    this._generateGanttAndStateData(
      taskDefinitions,
      scheduledTasks,
      actualScheduledTasks
    );
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
