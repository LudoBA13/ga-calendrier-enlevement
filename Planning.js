/**
* Stores planning data for a given year based on a range of dates.
* @param {number} year The year for the planning.
* @param {GoogleAppsScript.Spreadsheet.Range} range A 20x12 range containing dates.
* @returns {string[]} An array of 366 codes representing each day of the year.
*/
function storePlanning(year, range)
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

			result[dayIndex] = code;
		}
	}

	return result;
}

/**
* Stores the planning data into the 'Planning' sheet.
* @param {number} year The year to store.
* @param {GoogleAppsScript.Spreadsheet.Range} range The source range of dates.
*/
function savePlanning(year, range)
{
	const codes = storePlanning(year, range);
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName('Planning');

	if (!sheet)
	{
		sheet = ss.insertSheet('Planning');
	}

	const row = year - 2020;
	if (row < 1)
	{
		throw new Error('Year must be greater than 2020.');
	}

	// Check existing data to avoid unnecessary writes
	const existingRange = sheet.getRange(row, 1, 1, 367);
	const existingValues = existingRange.getValues()[0];
	const existingYear = existingValues[0];
	const existingCodes = existingValues.slice(1);

	// Check if year matches AND all codes match
	const yearMatches = (existingYear === year);
	const codesMatch = codes.every((code, index) => code === existingCodes[index]);

	if (yearMatches && codesMatch)
	{
		console.log('Planning for ' + year + ' is already up to date. Skipping write.');
		return;
	}

	// Write the year in column A
	sheet.getRange(row, 1).setValue(year);

	// Write the 366 codes in columns B to ... (366 columns starting from column 2)
	sheet.getRange(row, 2, 1, 366).setValues([codes]);
	console.log('Planning for ' + year + ' updated in the sheet.');
}
