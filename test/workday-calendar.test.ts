import { WorkdayCalendar } from '../src/workday-calendar';

describe('WorkdayCalendar', () => {
  const excludeSet = new Set([
    '2025/06/06',
    '2025/06/12',
    '2025/06/18',
    '2025/06/24',
    '2025/06/30',
  ]);
  const calendar = new WorkdayCalendar('2025/06/01', 30, (date, formatted) => {
    if (date.day() === 0 || date.day() === 6) return false;
    if (excludeSet.has(formatted)) return false;
    return true;
  });

  describe('getWorkdayAtOffset()', () => {
    it('test1', () => {
      expect(calendar.getWorkdayAtOffset('2025/06/01', 4)).toBe('2025/06/09');
    });
    it('test2', () => {
      expect(calendar.getWorkdayAtOffset('2025/06/02', 4)).toBe('2025/06/09');
    });
    it('test3', () => {
      expect(calendar.getWorkdayAtOffset('2025/06/26', -4)).toBe('2025/06/19');
    });
    it('test4', () => {
      expect(calendar.getWorkdayAtOffset('2025/06/27', -4)).toBe('2025/06/20');
    });
    it('test5', () => {
      expect(calendar.getWorkdayAtOffset('2025/06/29', -4)).toBe('2025/06/20');
    });
    it('test6', () => {
      expect(calendar.getWorkdayAtOffset('2025/06/30', -4)).toBe('2025/06/20');
    });
  });

  describe('getNextWorkday()', () => {
    it('test1', () => {
      expect(calendar.getNextWorkday('2025/06/01')).toBe('2025/06/02');
    });
    it('test2', () => {
      expect(calendar.getNextWorkday('2025/06/02')).toBe('2025/06/03');
    });
    it('test3', () => {
      expect(calendar.getNextWorkday('2025/06/05')).toBe('2025/06/09');
    });
    it('test4', () => {
      expect(calendar.getNextWorkday('2025/06/06')).toBe('2025/06/09');
    });
    it('test5', () => {
      expect(calendar.getNextWorkday('2025/06/11')).toBe('2025/06/13');
    });
  });

  describe('getPreviousWorkday()', () => {
    it('test1', () => {
      expect(calendar.getPreviousWorkday('2025/06/30')).toBe('2025/06/27');
    });
    it('test2', () => {
      expect(calendar.getPreviousWorkday('2025/06/29')).toBe('2025/06/27');
    });
    it('test3', () => {
      expect(calendar.getPreviousWorkday('2025/06/27')).toBe('2025/06/26');
    });
  });
});
