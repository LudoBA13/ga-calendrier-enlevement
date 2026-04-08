/**
* Extracts data from the 20x12 calendar range for a given year.
* @param {number} year The year for the planning.
* @param {GoogleAppsScript.Spreadsheet.Range} range A 20x12 range containing dates.
* @param {function(Date, string): any} transformFn A function that returns the value for each day of the year.
* @returns {any[]} An array of 366 values representing each day of the year.
*/
function extractDataFromCalendar(year, range, transformFn)
{
	const values = range.getValues();
	const result = new Array(366).fill('');
	const timeZone = Session.getScriptTimeZone();
	const dayCodes = ['Lu', 'Ma', 'Me', 'Je', 'Ve'];

	for (let r = 0; r < 20; r++)
	{
		const digit = Math.floor(r / 5) + 1;
		const dayCode = dayCodes[r % 5];
		const code = digit + dayCode;

		for (let c = 0; c < 12; c++)
		{
			const cellValue = values[r][c];

			if (!(cellValue instanceof Date) || cellValue.getFullYear() !== year)
			{
				continue;
			}

			// Use 'D' format for day in year (1-366). Map to 0-365 for index.
			const dayOfYear = Utilities.formatDate(cellValue, timeZone, 'D');
			const dayIndex = parseInt(dayOfYear) - 1;

			if (dayIndex < 0 || dayIndex >= 366)
			{
				continue;
			}

			result[dayIndex] = transformFn(cellValue, code);
		}
	}

	return result;
}

/**
* Stores planning data for a given year based on a range of dates.
* @param {number} year The year for the planning.
* @param {GoogleAppsScript.Spreadsheet.Range} range A 20x12 range containing dates.
* @returns {string[]} An array of 366 codes representing each day of the year.
*/
function storePlanning(year, range)
{
	return extractDataFromCalendar(year, range, (date, code) => code);
}

/**
* Extracts the month numbers for the days of the year present in the calendar.
* @param {number} year The year for the planning.
* @param {GoogleAppsScript.Spreadsheet.Range} range A 20x12 range containing dates.
* @returns {number[]} An array of 366 month numbers.
*/
function storeMonths(year, range)
{
	const dates = range.getValues()[0];
	dates[0] = new Date(year, 0, 1);
	const result = [];

	for (let i = 0; i < 11; i++)
	{
		const start = new Date(dates[i]);
		const end = new Date(dates[i + 1]);
		const diffInMs = end.getTime() - start.getTime();
		const diffInDays = Math.round(diffInMs / (1000 * 60 * 60 * 24));

		for (let d = 0; d < diffInDays; d++)
		{
			if (result.length < 366)
			{
				result.push(i + 1);
			}
		}
	}

	while (result.length < 366)
	{
		result.push(12);
	}

	return result;
}

/**
 * Saves data for a given year into a specific sheet.
 * @param {number} year The year to store.
 * @param {any[]} data The data to store.
 * @param {string} sheetName The name of the sheet.
 * @param {string} label Label for logging.
 */
function saveDataToSheet(year, data, sheetName, label)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(sheetName);

	if (!sheet)
	{
		sheet = ss.insertSheet(sheetName);
	}
	const row = year - 2020;
	if (row < 1)
	{
		throw new Error('Year must be greater than 2020.');
	}

	// Check existing data to avoid unnecessary writes
	const lastRow = sheet.getLastRow();
	if (row <= lastRow)
	{
		const existingRange = sheet.getRange(row, 1, 1, 367);
		const existingValues = existingRange.getValues()[0];

		// Check if year matches AND all data match
		if (existingValues[0] === year && JSON.stringify(data) === JSON.stringify(existingValues.slice(1)))
		{
			console.log(label + ' for ' + year + ' is already up to date. Skipping write.');
			return;
		}
	}

	// Write the year in column A
	sheet.getRange(row, 1).setValue(year);

	// Write the 366 codes in columns B to ... (366 columns starting from column 2)
	sheet.getRange(row, 2, 1, 366).setValues([data]);
	console.log(label + ' for ' + year + ' updated in the sheet.');
}

/**
 * Stores the planning data into the 'DateToPlanning' sheet.
 * @param {number} year The year to store.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The source range of dates.
 */
function savePlanning(year, range)
{
	const codes = storePlanning(year, range);
	saveDataToSheet(year, codes, 'DateToPlanning', 'Calendar codes');
}

/**
 * Stores the month data into the 'DateToPlanningMonth' sheet.
 * @param {number} year The year to store.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The source range of dates.
 */
function saveMonths(year, range)
{
	const months = storeMonths(year, range);
	saveDataToSheet(year, months, 'DateToPlanningMonth', 'Calendar months');
}
