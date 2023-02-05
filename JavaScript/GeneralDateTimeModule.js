// GeneralDateTimeModule.js
// All date-time related calculations

//  enumerators
const TimeMeasurementUnitTypeV000 = Object.freeze
({
	SECS_TMUT: Symbol(0),
	MINS_TMUT: Symbol(1),
	HOURS_TMUT: Symbol(2),
	DAYS_TMUT: Symbol(3),
	WEEKS_TMUT: Symbol(4),
	FORTNIGHTS_TMUT: Symbol(5),
	MONTHS_TMUT: Symbol(6),
	QTR_ANNUALS_TMUT: Symbol(7),
	SEMI_ANNUALS_TMUT: Symbol(8),
	YEARS_TMUT: Symbol(9),
	TWO_YEARS_TMUT: Symbol(10),
	THREE_YEARS_TMUT: Symbol(11),
	FOUR_YEARS_TMUT: Symbol(12),
	FIVE_YEARS_TMUT: Symbol(13),
	SIX_YEARS_TMUT: Symbol(14),
	SEVEN_YEARS_TMUT: Symbol(15),
	EIGHT_YEARS_TMUT: Symbol(16),
	NINE_YEARS_TMUT: Symbol(17),
	DECADES_TMUT: Symbol(18),
	TWO_DECADES_TMUT: Symbol(19),
	THREE_DECADES_TMUT: Symbol(20),
	FOUR_DECADES_TMUT: Symbol(21),
	FIVE_DECADES_TMUT: Symbol(22),
	CENTURIES_TMUT: Symbol(23),
	ONE_HUNDRED_AND_TWENTY_FIVE_YEARS_TMUT: Symbol(24),
	// 150 years should cover 1 x human lifetime
	ONE_HUNDRED_AND_FIFTY_YEARS_TMUT: Symbol(25),
})

//  global constants
const SECS_PER_MIN = 60;
const MINS_PER_HOUR = 60;
const HOURS_PER_DAY = 24;
const DAYS_PER_WEEK = 7;
// https://en.wikipedia.org/wiki/Year
const DAYS_PER_YEAR = 365.2425;
const WEEKS_PER_YEAR = (DAYS_PER_YEAR / DAYS_PER_WEEK);
const MONTHS_PER_YEAR = 12;
const WEEKS_PER_MONTH = (WEEKS_PER_YEAR / MONTHS_PER_YEAR);
//const ...

function 