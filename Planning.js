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
	const firstDayOfYear = new Date(year, 0, 1);
	const dayCodes = ['Lu', 'Ma', 'Me', 'Je', 'Ve'];

	for (let r = 0; r < 20; r++)
	{
		// Calculate the code for the current row
		// Digit 1-4: each digit covers 5 rows
		const digit = Math.floor(r / 5) + 1;
		// Day code rotation: Lu, Ma, Me, Je, Ve
		const dayCode = dayCodes[r % 5];
		const code = digit + dayCode;

		for (let c = 0; c < 12; c++)
		{
			const cellValue = values[r][c];

			if (cellValue instanceof Date)
			{
				// Ensure the date belongs to the correct year to avoid indexing errors
				if (cellValue.getFullYear() === year)
				{
					// Calculate day-of-year index (0-365)
					const diff = cellValue.getTime() - firstDayOfYear.getTime();
					const dayIndex = Math.floor(diff / (1000 * 60 * 60 * 24));

					if (dayIndex >= 0 && dayIndex < 366)
					{
						result[dayIndex] = code;
					}
				}
			}
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
	
	// Write the year in column A
	sheet.getRange(row, 1).setValue(year);
	
	// Write the 366 codes in columns B to ... (366 columns starting from column 2)
	// We use a 2D array [[]] for setValues
	sheet.getRange(row, 2, 1, 366).setValues([codes]);
}
