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
const DAYS_PER_MONTH = (DAYS_PER_YEAR / MONTHS_PER_YEAR);
const WEEKS_PER_MONTH = (WEEKS_PER_YEAR / MONTHS_PER_YEAR);
//const ...

// Define constructors for custom objects (similar to C structs and records)

// Define constructor for the complicated date time object - which has a use
//	case for budgeting
function defineDateTimeComplexObjV000(secs, mins, hours, days, weeks, fortnights, months, qtr_annuals, semi_annuals, years, two_years, three_years, four_years, five_years, six_years, seven_years, eight_years, nine_years, decades, two_decades, three_decades, four_decades, five_decades, centuries, one_hundred_and_twenty_five_years, one_hundred_and_fifty_years)
{
	this.secs = secs;
	this.mins = mins;
	this.hours = hours;
	this.days = days;
	this.weeks = weeks;
	this.fortnights = fortnights;
	this.months = months;
	this.qtr_annuals = qtr_annuals;
	this.semi_annuals = semi_annuals;
	this.years = years;
	this.two_years = two_years;
	this.three_years = three_years;
	this.four_years = four_years;
	this.five_years = five_years;
	this.six_years = six_years;
	this.seven_years = seven_years;
	this.eight_years = eight_years;
	this.nine_years = nine_years;
	this.decades = decades;
	this.two_decades = two_decades;
	this.three_decades = three_decades;
	this.four_decades = four_decades;
	this.five_decades = five_decades;
	this.centuries = centuries;
	this.one_hundred_and_twenty_five_years = one_hundred_and_twenty_five_years;
	this.one_hundred_and_fifty_years = one_hundred_and_fifty_years;
}

function defineDateTimeSimplerObjV000(secs, mins, hours, days, weeks, months, years, decades, centuries)
{
	this.secs = secs;
	this.mins = mins;
	this.hours = hours;
	this.days = days;
	this.weeks = weeks;
	this.months = months;
	this.years = years;
	this.decades = decades;
	this.centuries = centuries;
}

//	DateTimeSimpleObjV000
function defineDateTimeSimpleObjV000(secs, mins, hours, days, months, years)
{
	this.secs = secs;
	this.mins = mins;
	this.hours = hours;
	this.days = days;
	this.months = months;
	this.years = years;
}

//	@return Days Elapsed
function convertWeeksElapsedToDaysElapsedV000(weeks_elapsed)
{
	return (weeks_elapsed * DAYS_PER_WEEK);
}

//	@return Days Elapsed
function convertMonthsElapsedToDaysElapsedV000(months_elapsed)
{
	return (months_elapsed * DAYS_PER_MONTH);
}

//	@return Weeks Elapsed
function convertDaysElapsedToWeeksElapsedV000(days_elapsed)
{
	return (days_elapsed / DAYS_PER_WEEK);
}

//	@param DateTimeSimpleObjV000 type
function calcMinsElapsedViaDateTimeSimpleObjV000(date_time_elapsed_simple_obj)
{
	let result = 0;

	result = (date_time_elapsed_simple_obj.mins + (60 *
			(date_time_elapsed_simple_obj.hours + (24 *
			(date_time_elapsed_simple_obj.days + ((365 / 12) *
			(date_time_elapsed_simple_obj.months + (12 *
			(date_time_elapsed_simple_obj.years)))))))));

	return result;
}