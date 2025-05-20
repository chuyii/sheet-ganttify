export const COL_CHART = 21;
export const ROW_DATA = 5;

export const MONTH = [
  '1',
  '2',
  '3',
  '4',
  '5',
  '6',
  '7',
  '8',
  '9',
  '10',
  '11',
  '12',
];
export const WEEK = ['日', '月', '火', '水', '木', '金', '土'];

export const HOLIDAY_NATIONAL = '祝';
export const HOLIDAY_USER = 'ユ';

export const STATUS = {
  DONE: '完了',
  DELAYED: '遅延',
  IN_PROGRESS: '進行中',
  NEED_START: '要開始',
  NOT_STARTED: '未着手',
  NEED_REVIEW: '要見直',
} as const;

export type Status = (typeof STATUS)[keyof typeof STATUS];
