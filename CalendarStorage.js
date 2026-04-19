/**
 * Storage Format: DateToPlanning & TickToDate
 * ------------------------------------------
 * Both sheets follow a consistent bulk storage layout:
 * - Each row represents a full calendar year.
 * - Row number: (year - 2020).
 * - Column A: The year (e.g., 2026).
 * - Columns B to NC (366 columns): Sequential data for each day of the year (0-365).
 *
 * Data Types:
 * - DateToPlanning: Planning codes as strings (e.g., '1Lu', '2Ma', '3Me').
 * - TickToDate: Planning ticks as numbers (e.g., 2604110).
 *
 * Empty cells are represented by empty strings in the sheet and null in JavaScript structures.
 */

/**
 * Manages the storage of calendar data into spreadsheet sheets.
 */
class CalendarStorage
{
	/**
	 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss The spreadsheet instance.
	 */
	constructor(ss)
	{
		/** @private */
		this.ss = ss;
	}

	/**
	 * Refreshes the storage by retrieving data from all calendars and updating target sheets.
	 * @param {CalendarManager} calendarManager The manager to retrieve data from.
	 */
	refresh(calendarManager)
	{
		const calendarMap = calendarManager.getCalendarSheets();
		const planningDataMap = new Map;

		for (const [year, sheet] of calendarMap)
		{
			try
			{
				const range = calendarManager.getCalendarRange(sheet);
				planningDataMap.set(year, this._extractCalendarFromRange(range));
			}
			catch (error)
			{
				console.error('Erreur lors du traitement de ' + sheet.getName() + ' : ' + error.message);
			}
		}

		const tickDataMap = calendarManager.convertCalendarsToTicks(calendarMap);

		this._saveBulkDataToSheet(planningDataMap, 'DateToPlanning', 'Planning Codes');
		this._saveBulkDataToSheet(tickDataMap, 'TickToDate', 'Planning Ticks');
	}

	/**
	 * Extracts planning codes from the 20x12 calendar range.
	 * @param {GoogleAppsScript.Spreadsheet.Range} range A 20x12 range containing dates.
	 * @returns {any[]} An array of 366 values representing each day of the year.
	 * @private
	 */
	_extractCalendarFromRange(range)
	{
		const values = range.getValues();
		const result = new Array(366).fill(null);
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

				const dayIndex = this._getDayOfYear(cellValue);

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
	 * Saves multiple years of data to a specific sheet in bulk, pruning unchanged rows.
	 * @param {Map<number, any[]>} newDataMap A map of year to data array.
	 * @param {string} sheetName The name of the sheet.
	 * @param {string} label Label for logging.
	 * @private
	 */
	_saveBulkDataToSheet(newDataMap, sheetName, label)
	{
		let sheet = this.ss.getSheetByName(sheetName);

		if (!sheet)
		{
			sheet = this.ss.insertSheet(sheetName);
		}

		const years = Array.from(newDataMap.keys()).sort((a, b) => a - b);
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
			const rawData = newDataMap.get(year);
			const rowNum = year - 2020;

			if (rowNum < 1)
			{
				continue;
			}

			// Normalize input: convert nulls to empty strings for storage and comparison
			const data = rawData.map((val) => val === null ? '' : val);

			if (rowIdx < existingData.length)
			{
				const existingRow = existingData[rowIdx];
				if (existingRow[0] === year && JSON.stringify(data) === JSON.stringify(existingRow.slice(1)))
				{
					continue;
				}
			}

			// Write the year in column A
			sheet.getRange(rowNum, 1).setValue(year);
			// Write the 366 values in columns B to ...
			sheet.getRange(rowNum, 2, 1, 366).setValues([data]);
			updateCount++;
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
	 * Gets the day of the year (0-365) for a given date.
	 * @param {Date} date The date.
	 * @returns {number} The day of the year.
	 * @private
	 */
	_getDayOfYear(date)
	{
		const start = new Date(date.getFullYear(), 0, 0);
		const diff = date - start;
		const oneDay = 1000 * 60 * 60 * 24;
		return Math.floor(diff / oneDay) - 1;
	}
}