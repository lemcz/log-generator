#!usr/bin/env node

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
const fs = require('fs');

const LABEL = {
  DATE_FROM: 'LOG_DATE_FROM',
  DATE_TO: 'LOG_DATE_TO',
  CLIENT: 'LOG_CLIENT',
  ISSUE_NAME: 'LOG_ISSUE_NAME',
  PROJECT_HOURS: 'LOG_PROJECT_HOURS',
  INTERNAL_HOURS: 'LOG_INTERNAL_HOURS',
};

/**
 * CONFIGURATION
 */
const CONTRACTOR_ID = 'blemiec';
const FILE_FORMAT = 'xlsx';
const CLIENT = 'IGT';
const ISSUE_NAME = 'Aurora Navigator';
const WORK_HOURS_PER_DAY = 8;
const EXCEL_DATE_FORMAT = 'dd/MM/yyyy';

/**
 * READ JSON CONFIG
 */
const {
  contractor = CONTRACTOR_ID,
  fileFormat = FILE_FORMAT,
  client = CLIENT,
  projectName = ISSUE_NAME,
  workHours = WORK_HOURS_PER_DAY,
  dateFormat = EXCEL_DATE_FORMAT,
} = JSON.parse(fs.readFileSync('./config.json').toString());

const currentDate = new Date();
const parseYear = argv['y'] || currentDate.getFullYear();
const parseMonth = argv['m'] || currentDate.getMonth() + 1;
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
    [LABEL.DATE_FROM]: formatDate(dateFrom, dateFormat),
    [LABEL.DATE_TO]: formatDate(dateTo, dateFormat),
    [LABEL.INTERNAL_HOURS]: 0,
    [LABEL.PROJECT_HOURS]: differenceInBusinessDays(addDays(dateTo, 1), dateFrom) * workHours,
    [LABEL.CLIENT]: client,
    [LABEL.ISSUE_NAME]: projectName,
  };
});

const workbook = XLSX.utils.book_new();
const worksheet = XLSX.utils.json_to_sheet(data, {
  header: [LABEL.DATE_FROM, LABEL.DATE_TO, LABEL.CLIENT, LABEL.ISSUE_NAME, LABEL.PROJECT_HOURS, LABEL.INTERNAL_HOURS],
});
const projectHoursSumCell = XLSX.utils.encode_cell({ c: 4, r: 10 });
const internalHoursSumCell = XLSX.utils.encode_cell({ c: 5, r: 10 });
worksheet['!ref'] = XLSX.utils.encode_range({ s: { c: 0, r: 0 }, e: { c: 11, r: 11 } });
worksheet[projectHoursSumCell] = { f: 'SUM(E2:E7)' };
worksheet[internalHoursSumCell] = { f: 'SUM(F2:F7)' };
XLSX.utils.book_append_sheet(workbook, worksheet);

writeFile(workbook, {
  fileName: `${contractor}_${format(logReportDate, 'yyyyMM')}`,
  fileFormat: fileFormat,
});

function writeFile(workBook, { fileName, fileFormat }) {
  XLSX.writeFile(workbook, `${fileName}.${fileFormat}`);
}

function formatDate(date, { dateFormat = 'dd/MM/yyyy', weekStart = 1 } = {}) {
  return format(date, dateFormat, weekStart);
}

function getISOStringWithoutTime(date = new Date()) {
  return date.toISOString().split('T')[0];
}
