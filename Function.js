const yearPlanningCache = {};
const dayPlanningCache = {};

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
