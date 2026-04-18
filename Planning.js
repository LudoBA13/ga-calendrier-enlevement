/**
 * Extracts dates from the 20x12 calendar range.
 * @param {GoogleAppsScript.Spreadsheet.Range} range A 20x12 range containing dates.
 * @returns {any[]} An array of 366 values representing each day of the year.
 */
function extractCalendarFromRange(range)
{
	const values = range.getValues();
	const result = new Array(366).fill('');
	const dayCodes = ['Lu', 'Ma', 'Me', 'Je', 'Ve'];

	for (let r = 0; r < 20; r++)
	{
		const digit = Math.floor(r / 5) + 1;
		const dayCode = dayCodes[r % 5];
		const code = digit + dayCode;

		for (let c = 0; c < 12; c++)
		{
			const cellValue = values[r][c];

			if (!(cellValue instanceof Date))
			{
				continue;
			}

			// Use pure JS for day in year (0-365).
			const dayIndex = getDayOfYear(cellValue);

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
 * Stores planning data for a given year based on a range of dates.
 * @param {number} year The year for the planning.
 * @param {GoogleAppsScript.Spreadsheet.Range} range A 20x12 range containing dates.
 * @returns {string[]} An array of 366 codes representing each day of the year.
 */
function storePlanning(year, range)
{
	return extractCalendarFromRange(range);
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
 * Finds the 20x12 calendar range in the given sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
 * @returns {GoogleAppsScript.Spreadsheet.Range} The range.
 * @throws {Error} If the range is not found.
 */
function getCalendarRange(sheet)
{
	// Dynamically locate the first row in column A that matches /^1.*lundi$/
	const lastRow = sheet.getLastRow();
	if (lastRow < 1)
	{
		throw new Error('La feuille est vide ou ne peut pas être lue.');
	}
	const colAValues = sheet.getRange(1, 1, lastRow).getDisplayValues();
	let startRow = -1;

	for (let i = 0; i < lastRow; i++)
	{
		if (/^1.*lundi$/.test(colAValues[i][0]))
		{
			startRow = i + 1;
			break;
		}
	}

	if (startRow === -1)
	{
		throw new Error('Impossible de trouver le début du planning (une cellule en colonne A commençant par "1" et finissant par "lundi").');
	}

	// Range starting at Column B of that row, 20 rows by 12 columns
	return sheet.getRange(startRow, 2, 20, 12);
}

/**
 * Synchronizes all caches by processing all 'CalendrierX' sheets.
 */
function syncAllCaches()
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheets = ss.getSheets();
	const planningData = {};
	const monthData = {};

	sheets.forEach((sheet) =>
	{
		const sheetName = sheet.getName();
		const match = sheetName.match(/^Calendrier(20\d+)$/);
		if (match)
		{
			const year = parseInt(match[1]);
			try
			{
				const range = getCalendarRange(sheet);
				planningData[year] = storePlanning(year, range);
				monthData[year] = storeMonths(year, range);
			}
			catch (error)
			{
				console.error('Erreur lors du traitement de ' + sheetName + ' : ' + error.message);
			}
		}
	});

	saveBulkDataToSheet(planningData, 'DateToPlanning', 'Calendar codes');
	saveBulkDataToSheet(monthData, 'DateToPlanningMonth', 'Calendar months');
}

/**
 * Saves multiple years of data to a specific sheet in bulk, pruning unchanged rows.
 * @param {Object.<number, any[]>} newDataMap A map of year to data array.
 * @param {string} sheetName The name of the sheet.
 * @param {string} label Label for logging.
 */
function saveBulkDataToSheet(newDataMap, sheetName, label)
{
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(sheetName);

	if (!sheet)
	{
		sheet = ss.insertSheet(sheetName);
	}

	const years = Object.keys(newDataMap).map(Number).sort((a, b) => a - b);
	if (years.length === 0)
	{
		return;
	}

	const maxYear = years[years.length - 1];
	const expectedLastRow = maxYear - 2020;

	// Ensure sheet has enough rows
	const currentMaxRows = sheet.getMaxRows();
	if (currentMaxRows < expectedLastRow)
	{
		sheet.insertRowsAfter(currentMaxRows, expectedLastRow - currentMaxRows);
	}

	// Read existing data
	const lastRow = sheet.getLastRow();
	const existingData = lastRow > 0 ? sheet.getRange(1, 1, lastRow, 367).getValues() : [];

	let updateCount = 0;

	for (const year of years)
	{
		const rowIdx = year - 2020 - 1; // 0-based index for existingData array
		const data = newDataMap[year];
		const rowNum = year - 2020;

		if (rowNum < 1)
		{
			continue;
		}

		let shouldUpdate = true;

		if (rowIdx < existingData.length)
		{
			const existingRow = existingData[rowIdx];
			if (existingRow[0] === year && JSON.stringify(data) === JSON.stringify(existingRow.slice(1)))
			{
				shouldUpdate = false;
			}
		}

		if (shouldUpdate)
		{
			// Write the year in column A
			sheet.getRange(rowNum, 1).setValue(year);
			// Write the 366 codes in columns B to ... (366 columns starting from column 2)
			sheet.getRange(rowNum, 2, 1, 366).setValues([data]);
			updateCount++;
		}
	}

	if (updateCount > 0)
	{
		console.log(label + ': ' + updateCount + ' years updated.');
	}
	else
	{
		console.log(label + ': All up to date. Skipping write.');
	}
}

/**
 * Stores the planning data into the 'DateToPlanning' sheet.
 * @param {number} year The year to store.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The source range of dates.
 */
function savePlanning(year, range)
{
	const codes = storePlanning(year, range);
	const dataMap = {};
	dataMap[year] = codes;
	saveBulkDataToSheet(dataMap, 'DateToPlanning', 'Calendar codes');
}

/**
 * Stores the month data into the 'DateToPlanningMonth' sheet.
 * @param {number} year The year to store.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The source range of dates.
 */
function saveMonths(year, range)
{
	const months = storeMonths(year, range);
	const dataMap = {};
	dataMap[year] = months;
	saveBulkDataToSheet(dataMap, 'DateToPlanningMonth', 'Calendar months');
}
