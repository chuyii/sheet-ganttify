import dayjs from 'dayjs';

/**
 * A predicate function that determines whether a given date is a workday.
 * @param date A Dayjs object representing the current date being checked.
 * @param formatted The same date formatted as "YYYY/MM/DD".
 * @returns True if the date is considered a workday, otherwise false.
 */
export type WorkdayPredicate = (
  date: dayjs.Dayjs,
  formatted: string
) => boolean;

/**
 * A calendar utility class for managing and calculating dates based on workdays (i.e., business days).
 */
export class WorkdayCalendar {
  private workdayList: string[] = [];
  private dateToIndexMap = new Map<string, number>();

  /**
   * Constructs a workday calendar starting from a specific date, over a given number of calendar days.
   * @param startDate The starting date in "YYYY/MM/DD" format.
   * @param totalDays The total number of calendar days to include in the calendar.
   * @param isWorkday Optional function to determine whether a date is a workday. Defaults to treating all dates as workdays.
   */
  constructor(
    startDate: string,
    totalDays: number,
    isWorkday: WorkdayPredicate = () => true
  ) {
    let currentDate = dayjs(startDate);
    let index = 0;

    for (let i = 0; i < totalDays; i++) {
      const formatted = currentDate.format('YYYY/MM/DD');
      if (isWorkday(currentDate, formatted)) {
        this.dateToIndexMap.set(formatted, index++);
        this.workdayList.push(formatted);
      }
      currentDate = currentDate.add(1, 'day');
    }
  }

  /**
   * Returns the workday string that is N workdays offset from the specified base date.
   * If the base date is not a workday, the nearest valid workday is used depending on the direction.
   * @param baseDate The base date in "YYYY/MM/DD" format.
   * @param offset The number of workdays to move (positive = forward, negative = backward).
   * @returns The resulting workday string, or null if not resolvable.
   */
  public getWorkdayAtOffset(baseDate: string, offset: number): string | null {
    let baseIndex = this.dateToIndexMap.get(baseDate);

    // If baseDate is not a workday, find nearest workday in the direction of offset
    if (baseIndex === undefined) {
      const nearestIndex = this.findNearestWorkdayIndex(baseDate, offset);
      if (nearestIndex === null) return null;
      baseIndex = nearestIndex;
    }

    const targetIndex = baseIndex + offset;
    if (targetIndex < 0 || targetIndex >= this.workdayList.length) return null;
    return this.workdayList[targetIndex];
  }

  /**
   * Finds the nearest workday index based on the direction of offset.
   * @param baseDate The base date in "YYYY/MM/DD" format.
   * @param offset Direction to search: positive for forward, negative for backward.
   * @returns Index of the nearest workday or null if none found.
   */
  private findNearestWorkdayIndex(
    baseDate: string,
    offset: number
  ): number | null {
    const list = this.workdayList;
    let low = 0;
    let high = list.length - 1;

    while (low <= high) {
      const mid = Math.floor((low + high) / 2);
      const midVal = list[mid];

      if (midVal === baseDate) {
        return mid; // unlikely since baseDate is not in dateToIndex, but just in case
      } else if (midVal < baseDate) {
        low = mid + 1;
      } else {
        high = mid - 1;
      }
    }

    if (offset > 0) {
      // next workday after baseDate
      return low < list.length ? low : null;
    } else if (offset < 0) {
      // previous workday before baseDate
      return high >= 0 ? high : null;
    } else {
      // offset === 0 and not a workday
      return null;
    }
  }

  /**
   * Returns the closest workday that is after the specified base date.
   * @param baseDate A date string in "YYYY/MM/DD" format.
   * @returns The next workday, or null if none found.
   */
  public getNextWorkday(baseDate: string): string | null {
    if (this.dateToIndexMap.has(baseDate))
      return this.getWorkdayAtOffset(baseDate, +1);

    const index = this.findNearestWorkdayIndex(baseDate, +1);
    return index !== null ? this.workdayList[index] : null;
  }

  /**
   * Returns the closest workday that is before the specified base date.
   * @param baseDate A date string in "YYYY/MM/DD" format.
   * @returns The previous workday, or null if none found.
   */
  public getPreviousWorkday(baseDate: string): string | null {
    if (this.dateToIndexMap.has(baseDate))
      return this.getWorkdayAtOffset(baseDate, -1);

    const index = this.findNearestWorkdayIndex(baseDate, -1);
    return index !== null ? this.workdayList[index] : null;
  }
}
