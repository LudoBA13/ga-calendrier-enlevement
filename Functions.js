const yearPlanningCache = {};
const yearMonthCache = {};
const dayPlanningCache = {};
const dayMonthCache = {};
const planningDateCache = {};

/**
 * Gets the planning map for a specific year from the 'DateToPlanning' sheet.
 * @param {number} year The year to retrieve.
 * @returns {Array<string>} The values from columns B to NC for the given year.
 */
function getDateToPlanningMap(year)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('DateToPlanning');

	if (!sheet)
	{
		throw new Error('Sheet "DateToPlanning" not found.');
	}

	const row = year - 2020;
	if (row < 1)
	{
		throw new Error('Year must be greater than 2020.');
	}

	// Columns B to NC is 366 columns starting from column 2.
	const range = sheet.getRange(row, 2, 1, 366);
	return range.getValues()[0];
}

/**
 * Gets the month map for a specific year from the 'DateToPlanningMonth' sheet.
 * @param {number} year The year to retrieve.
 * @returns {Array<number|string>} The values from columns B to NC for the given year.
 */
function getDateToPlanningMonthMap(year)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('DateToPlanningMonth');

	if (!sheet)
	{
		throw new Error('Sheet "DateToPlanningMonth" not found.');
	}

	const row = year - 2020;
	if (row < 1)
	{
		throw new Error('Year must be greater than 2020.');
	}

	// Columns B to NC is 366 columns starting from column 2.
	const range = sheet.getRange(row, 2, 1, 366);
	return range.getValues()[0];
}

/**
 * Gets the planning map for a specific year from the 'PlanningToDate' sheet.
 * @param {number} year The year to retrieve.
 * @returns {Array<Array<string>>} The values from columns B to IG for the given year, split into 12 months.
 */
function getPlanningToDateMap(year)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getSheetByName('PlanningToDate');

	if (!sheet)
	{
		throw new Error('Sheet "PlanningToDate" not found.');
	}

	const row = year - 2020;
	if (row < 1)
	{
		throw new Error('Year must be greater than 2020.');
	}

	// Retrieve 20 planning days multiplied by 12 month
	const range = sheet.getRange(row, 2, 1, 240);
	const values = range.getValues()[0];
	const chunks = [[]];
	for (let i = 0; i < 12; i++)
	{
		chunks.push(values.slice(i * 20, (i + 1) * 20));
	}
	return chunks;
}

/**
 * Custom function to get dates for a specific planning code.
 * @param {string} code The planning code (e.g., '1Lu').
 * @param {number} year The year.
 * @param {number} month The month (1-12).
 * @returns {Date|string} The date or an empty string.
 * @customfunction
 */
function PLANNING_TO_DATE(code, year, month)
{
	if (!code || typeof code !== 'string' || code.length < 3)
	{
		return '';
	}

	const week = parseInt(code.substring(0, 1));
	const dayCode = code.substring(1, 3);
	const dayCodes = ['Lu', 'Ma', 'Me', 'Je', 'Ve'];
	const dayIndex = dayCodes.indexOf(dayCode);

	if (isNaN(week) || dayIndex === -1)
	{
		return '';
	}

	const index = (week - 1) * 5 + dayIndex;
	const cacheKey = year + '-' + month;

	if (!planningDateCache[cacheKey])
	{
		const yearMap = getPlanningToDateMap(year);
		if (!yearMap[month])
		{
			return '';
		}
		planningDateCache[cacheKey] = yearMap[month];
	}

	const monthDays = planningDateCache[cacheKey];
	const date = monthDays[index];

	return date instanceof Date ? date : '';
}

/**
 * Custom function to get planning code for a date or range of dates.
 * @param {Date|Array<Date>} input The date or range of dates.
 * @returns {string|Array<string>} The planning code or array of codes.
 * @customfunction
 */
function DATE_TO_PLANNING(input)
{
	if (Array.isArray(input))
	{
		return input.map(function(row)
		{
			return row.map(function(cell)
			{
				return cell instanceof Date ? dateToPlanning(cell) : '';
			});
		});
	}

	return input instanceof Date ? dateToPlanning(input) : '';
}

/**
 * Custom function to get month number for a date or range of dates.
 * @param {Date|Array<Date>} input The date or range of dates.
 * @returns {number|Array<number>} The month number or array of month numbers.
 * @customfunction
 */
function DATE_TO_MONTH_NUM(input)
{
	if (Array.isArray(input))
	{
		return input.map(function(row)
		{
			return row.map(function(cell)
			{
				return cell instanceof Date ? dateToMonthNum(cell) : '';
			});
		});
	}

	return input instanceof Date ? dateToMonthNum(input) : '';
}

/**
 * Gets the month number for a specific date from the 'DateToPlanningMonth' sheet.
 * @param {Date} date The date.
 * @returns {number|string} The month number or empty string.
 */
function dateToMonthNum(date)
{
	const time = date.getTime();
	if (dayMonthCache[time])
	{
		return dayMonthCache[time];
	}

	const year = date.getFullYear();
	if (!yearMonthCache[year])
	{
		yearMonthCache[year] = getDateToPlanningMonthMap(year);
	}

	const dayOfYear = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'D')) - 1;
	const monthNum = yearMonthCache[year][dayOfYear];

	dayMonthCache[time] = monthNum;
	return monthNum;
}

/**
 * Gets the planning code for a specific date.
 * @param {Date} date The date.
 * @returns {string} The planning code.
 */
function dateToPlanning(date)
{
	const time = date.getTime();
	if (dayPlanningCache[time])
	{
		return dayPlanningCache[time];
	}

	const year = date.getFullYear();
	if (!yearPlanningCache[year])
	{
		yearPlanningCache[year] = getDateToPlanningMap(year);
	}

	const dayOfYear = parseInt(Utilities.formatDate(date, Session.getScriptTimeZone(), 'D')) - 1;
	const code = yearPlanningCache[year][dayOfYear];

	dayPlanningCache[time] = code;
	return code;
}
