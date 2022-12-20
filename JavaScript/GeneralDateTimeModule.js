// GeneralDateTimeModule.js
// All date-time related calculations

const SECS_PER_MIN = 60;
const MINS_PER_HOUR = 60;
const HOURS_PER_DAY = 24;
const DAYS_PER_WEEK = 7;
// https://en.wikipedia.org/wiki/Year
const DAYS_PER_YEAR = 365.2425;
const WEEKS_PER_YEAR = (DAYS_PER_YEAR / DAYS_PER_WEEK);
const MONTHS_PER_YEAR = 12;
const WEEKS_PER_MONTH = (WEEKS_PER_YEAR / MONTHS_PER_YEAR);