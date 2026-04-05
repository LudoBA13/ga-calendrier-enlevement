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

	console.log('Starting storePlanning for year ' + year);

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

			if (r === 0 || c === 0 || dayIndex === 0)
			{
				console.log('Debug: Cell[' + r + '][' + c + '] = ' + cellValue + ', dayOfYear="' + dayOfYear + '", dayIndex=' + dayIndex + ', code=' + code);
			}

			if (dayIndex < 0 || dayIndex >= 366)
			{
				continue;
			}

			if (result[dayIndex] === '')
			{
				result[dayIndex] = code;
			}
			else
			{
				console.log('Warning: Duplicate date found at index ' + dayIndex + '. Existing: ' + result[dayIndex] + ', New: ' + code);
			}
		}
	}

	const filledCount = result.filter(v => v !== '').length;
	console.log('Planning stored. Total days filled: ' + filledCount);
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
	
	// Write the year in column A
	sheet.getRange(row, 1).setValue(year);
	
	// Write the 366 codes in columns B to ... (366 columns starting from column 2)
	// We use a 2D array [[]] for setValues
	sheet.getRange(row, 2, 1, 366).setValues([codes]);
}
