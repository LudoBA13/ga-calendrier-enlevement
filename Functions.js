/**
 * Tick Encoding Scheme: YYMMWDT
 * ----------------------------
 * A "tick" is a numerical representation of a specific planning day:
 * - YY: Year (last two digits, e.g., 26 for 2026) -> Multiplier 100,000
 * - MM: Month (1-12) -> Multiplier 1,000
 * - W:  Week (1-4) -> Multiplier 100
 * - D:  Day (1:Lu, 2:Ma, 3:Me, 4:Je, 5:Ve) -> Multiplier 10
 * - T:  Timeslot (1:Md, 2:Mf, 3:Ap)
 *
 * Example: 2604110 represents 2026, April, Week 1, Monday (Lu).
 */

const yearPlanningCache = {};
const yearMonthCache = {};
const dayPlanningCache = {};
const dayMonthCache = {};
const planningDateCache = {};
const yearPlanningDateMapCache = {};
const planningIdx = {"1Lu":0,"1Ma":1,"1Me":2,"1Je":3,"1Ve":4,"2Lu":5,"2Ma":6,"2Me":7,"2Je":8,"2Ve":9,"3Lu":10,"3Ma":11,"3Me":12,"3Je":13,"3Ve":14,"4Lu":15,"4Ma":16,"4Me":17,"4Je":18,"4Ve":19};
const planningToTick = {"1Lu":110,"1Ma":120,"1Me":130,"1Je":140,"1Ve":150,"2Lu":210,"2Ma":220,"2Me":230,"2Je":240,"2Ve":250,"3Lu":310,"3Ma":320,"3Me":330,"3Je":340,"3Ve":350,"4Lu":410,"4Ma":420,"4Me":430,"4Je":440,"4Ve":450};

/**
 * Ensures a sheet exists, is correctly sized, and contains an IMPORTRANGE formula.
 * @param {string} sheetName The name of the sheet.
 * @param {number} expectedColumns The number of columns the sheet should have.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet.
 */
function ensureSheet(sheetName, expectedColumns)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(sheetName);
	if (!sheet)
	{
		sheet = ss.insertSheet(sheetName);
		const masterId = '1H3WpuEOI8mJvXh1kLyX2Cn4Rwzt8YnRa07-zWHOJBuA';
		const lastColLetter = sheet.getRange(1, expectedColumns).getA1Notation().replace(/\d/g, '');
		sheet.getRange(1, 1).setFormula('=IMPORTRANGE("' + masterId + '"; "' + sheetName + '!A:' + lastColLetter + '")');

		// Resize columns
		const currentCols = sheet.getMaxColumns();
		if (currentCols > expectedColumns)
		{
			sheet.deleteColumns(expectedColumns + 1, currentCols - expectedColumns);
		}
		else if (currentCols < expectedColumns)
		{
			sheet.insertColumnsAfter(currentCols, expectedColumns - currentCols);
		}

		// Resize rows (ensure at least 50 rows)
		const expectedRows = 50;
		const currentRows = sheet.getMaxRows();
		if (currentRows > expectedRows)
		{
			sheet.deleteRows(expectedRows + 1, currentRows - expectedRows);
		}
		else if (currentRows < expectedRows)
		{
			sheet.insertRowsAfter(currentRows, expectedRows - currentRows);
		}
	}
	return sheet;
}

/**
 * Gets the planning map for a specific year from the 'DateToPlanning' sheet.
 * @param {number} year The year to retrieve.
 * @returns {Array<string>} The values from columns B to NC for the given year.
 */
function getDateToPlanningMap(year)
{
	const sheet = ensureSheet('DateToPlanning', 367);

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
	const sheet = ensureSheet('DateToPlanningMonth', 367);

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
	if (yearPlanningDateMapCache[year])
	{
		return yearPlanningDateMapCache[year];
	}

	const sheet = ensureSheet('PlanningToDate', 241);

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

	yearPlanningDateMapCache[year] = chunks;
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
	if (typeof code !== 'string')
	{
		return '';
	}

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

	const index = planningIdx[code];
	const date = planningDateCache[cacheKey][index];

	return date instanceof Date ? date : '';
}

/**
 * Custom function to get the tick for a date or range of dates.
 * @param {Date|Array<Date>} input The date or range of dates.
 * @returns {number|Array<number>} The tick or array of ticks.
 * @customfunction
 */
function DATE_TO_TICK(input)
{
	if (Array.isArray(input))
	{
		return input.map(function(row)
		{
			return row.map(function(cell)
			{
				return cell instanceof Date ? dateToTick(cell) : '';
			});
		});
	}

	return input instanceof Date ? dateToTick(input) : '';
}

/**
 * Custom function to get the date for a tick or range of ticks.
 * @param {number|Array<number>} input The tick or range of ticks.
 * @returns {Date|Array<Date>|string} The date, array of dates, or empty string.
 * @customfunction
 */
function TICK_TO_DATE(input)
{
	if (Array.isArray(input))
	{
		return input.map(function(row)
		{
			return row.map(function(cell)
			{
				if (cell === '' || cell === null || cell === undefined)
				{
					return '';
				}
				const tick = parseInt(cell);
				return !isNaN(tick) ? tickToDate(tick) : '';
			});
		});
	}

	if (input === '' || input === null || input === undefined)
	{
		return '';
	}
	const tick = parseInt(input);
	return !isNaN(tick) ? tickToDate(tick) : '';
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

/**
 * Gets the tick for a specific date.
 * @param {Date} date The date.
 * @returns {number|string} The tick or empty string.
 */
function dateToTick(date)
{
	const year = date.getFullYear() % 100;
	const month = dateToMonthNum(date) || 0;
	const planning = dateToPlanning(date);
	const tickValue = planningToTick[planning] || 0;

	return year * 100000 + month * 1000 + tickValue;
}

/**
 * Gets the date for a specific tick.
 * @param {number} tick The tick.
 * @returns {Date|string} The date or empty string.
 */
function tickToDate(tick)
{
	if (!tick || isNaN(tick))
	{
		throw new Error('Invalid tick value: ' + tick);
	}

	const year = 2000 + Math.floor(tick / 100000);
	const month = Math.floor((tick % 100000) / 1000);
	const tickValue = tick % 1000;

	const week = Math.floor(tickValue / 100);
	const dayDigit = Math.floor((tickValue % 100) / 10);
	const dayCode = { 1: 'Lu', 2: 'Ma', 3: 'Me', 4: 'Je', 5: 'Ve' }[dayDigit] || '??';

	if (dayCode === '??')
	{
		throw new Error('Invalid day digit in tick: ' + tick);
	}

	const code = week + dayCode;

	return PLANNING_TO_DATE(code, year, month);
}
