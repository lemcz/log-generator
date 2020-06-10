const XLSX = require('xlsx');
const startOfMonth = require('date-fns/startOfMonth');
const endOfWeek = require('date-fns/endOfWeek');
const endOfMonth = require('date-fns/endOfMonth');
const eachWeekOfInterval = require('date-fns/eachWeekOfInterval');
const differenceInBusinessDays = require('date-fns/differenceInBusinessDays');
const addDays = require('date-fns/addDays');
const isAfter = require('date-fns/isAfter');
const parseISO = require('date-fns/parseISO');
const format = require('date-fns/format');
const argv = require('minimist')(process.argv.slice(2));

const LABEL = {
  DATE_FROM: 'LOG_DATE_FROM',
  DATE_TO: 'LOG_DATE_TO',
  CLIENT: 'LOG_CLIENT',
  ISSUE_NAME: 'LOG_ISSUE_NAME',
  PROJECT_HOURS: 'LOG_PROJECT_HOURS',
  INTERNAL_HOURS: 'LOG_INTERNAL_HOURS',
};

const CONTRACTOR_ID = 'blemiec';
const FILE_FORMAT = 'xlsx';
const CLIENT = 'IGT';
const ISSUE_NAME = 'Aurora Navigator';
const WORK_HOURS_PER_DAY = 8;
const EXCEL_DATE_FORMAT = 'dd/MM/yyyy';

const currentDate = new Date();
const parseYear = argv['y'] || currentDate.getFullYear();
const parseMonth = argv['m'] || currentDate.getMonth();
const dateToParse = new Date(parseYear, parseMonth);
const logReportDate = parseISO(getISOStringWithoutTime(dateToParse));
const month = {
  start: startOfMonth(logReportDate),
  end: endOfMonth(logReportDate),
};

const weeks = eachWeekOfInterval(
  {
    start: month.start,
    end: month.end,
  },
  { weekStartsOn: 1 },
);

const data = weeks.map((weekStart) => {
  const endDay = endOfWeek(weekStart, { weekStartsOn: 1 });
  const dateFrom = isAfter(month.start, weekStart) ? month.start : weekStart;
  const dateTo = isAfter(endDay, month.end) ? month.end : endDay;
  return {
    [LABEL.DATE_FROM]: format(dateFrom, EXCEL_DATE_FORMAT, { weekStartsOn: 1 }),
    [LABEL.DATE_TO]: format(dateTo, EXCEL_DATE_FORMAT, { weekStartsOn: 1 }),
    [LABEL.INTERNAL_HOURS]: 0,
    [LABEL.PROJECT_HOURS]: differenceInBusinessDays(addDays(dateTo, 1), dateFrom) * WORK_HOURS_PER_DAY,
    [LABEL.CLIENT]: CLIENT,
    [LABEL.ISSUE_NAME]: ISSUE_NAME,
  };
});

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(data, {
  header: [LABEL.DATE_FROM, LABEL.DATE_TO, LABEL.CLIENT, LABEL.ISSUE_NAME, LABEL.PROJECT_HOURS, LABEL.INTERNAL_HOURS],
});
const cellRef = XLSX.utils.encode_cell({ c: 4, r: 10 });
const cellRef2 = XLSX.utils.encode_cell({ c: 5, r: 10 });
worksheet['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 11, r: 11 } });
worksheet[cellRef] = { f: 'SUM(E2:E6)' };
worksheet[cellRef2] = { f: 'SUM(F2:F6)' };
XLSX.utils.book_append_sheet(workbook, worksheet);

XLSX.writeFile(workbook, `${CONTRACTOR_ID}_${format(logReportDate, 'yyyyMM')}.${FILE_FORMAT}`);

function getISOStringWithoutTime(date = new Date()) {
  return date.toISOString().split('T')[0];
}
