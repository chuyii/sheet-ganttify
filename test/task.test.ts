import { TaskDefinition, resolveSchedule } from '../src/task';
import { WorkdayCalendar } from '../src/workday-calendar';

type TestData = [
  [
    TaskDefinition['id'],
    TaskDefinition['startDate'],
    TaskDefinition['endDate'],
    TaskDefinition['duration'],
    TaskDefinition['id'][], // startsAfter
    TaskDefinition['id'][], // endsBefore
  ],
  [
    NonNullable<TaskDefinition['startDate']>,
    NonNullable<TaskDefinition['endDate']>,
  ],
][];

describe('task', () => {
  describe('resolveSchedule()', () => {
    it('test1', () => {
      const calendar = new WorkdayCalendar('2024/07/01', 365);
      const data: TestData = [
        [
          [0, '2025/01/01', null, 3, [], []],
          ['2025/01/01', '2025/01/03'],
        ],
        [
          [1, null, null, 3, [], [0]],
          ['2024/12/29', '2024/12/31'],
        ],
        [
          [2, null, null, 1, [], [1]],
          ['2024/12/28', '2024/12/28'],
        ],
        [
          [3, null, null, 1, [0], []],
          ['2025/01/04', '2025/01/04'],
        ],
      ];
      const tasks: TaskDefinition[] = data.map(d => ({
        id: d[0][0],
        startDate: d[0][1],
        endDate: d[0][2],
        duration: d[0][3],
        startsAfter: new Set(d[0][4]),
        endsBefore: new Set(d[0][5]),
      }));
      expect(resolveSchedule(tasks, calendar)[0]).toStrictEqual(
        tasks.map((task, index) => ({
          ...task,
          startDate: data[index][1][0],
          endDate: data[index][1][1],
        }))
      );
    });

    it('test2', () => {
      const calendar = new WorkdayCalendar('2024/07/01', 365, date => {
        if (date.day() === 0 || date.day() === 6) return false;
        return true;
      });
      const data: TestData = [
        [
          [0, '2025/01/01', null, 3, [], []],
          ['2025/01/01', '2025/01/03'],
        ],
        [
          [1, null, null, 3, [], [0]],
          ['2024/12/27', '2024/12/31'],
        ],
        [
          [2, null, null, 1, [], [1]],
          ['2024/12/26', '2024/12/26'],
        ],
        [
          [3, null, null, 1, [0], []],
          ['2025/01/06', '2025/01/06'],
        ],
      ];
      const tasks: TaskDefinition[] = data.map(d => ({
        id: d[0][0],
        startDate: d[0][1],
        endDate: d[0][2],
        duration: d[0][3],
        startsAfter: new Set(d[0][4]),
        endsBefore: new Set(d[0][5]),
      }));
      expect(resolveSchedule(tasks, calendar)[0]).toStrictEqual(
        tasks.map((task, index) => ({
          ...task,
          startDate: data[index][1][0],
          endDate: data[index][1][1],
        }))
      );
    });
  });
});
