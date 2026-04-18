/**
 * Manages calendar synchronization and data retrieval from a spreadsheet.
 */
class CalendarManager
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
	 * Scans all sheets and returns those matching the "CalendrierYYYY" pattern.
	 * @returns {Map<number, GoogleAppsScript.Spreadsheet.Sheet>} Map where key is the year and value is the sheet.
	 */
	getCalendarSheets()
	{
		const sheets = this.ss.getSheets();
		const calendars = [];
		const regex = /^Calendrier(20[0-9]{2})$/;

		for (const sheet of sheets)
		{
			const sheetName = sheet.getName();
			const yearMatch = sheetName.match(regex);

			if (yearMatch)
			{
				const year = Number(yearMatch[1]);
				calendars.push({ year: year, sheet: sheet });
			}
		}

		// Sort by year ascending
		calendars.sort((a, b) => a.year - b.year);

		const calendarMap = new Map;
		for (const cal of calendars)
		{
			calendarMap.set(cal.year, cal.sheet);
		}

		return calendarMap;
	}

	/**
	 * Finds the 20x12 calendar range in the given sheet.
	 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
	 * @returns {GoogleAppsScript.Spreadsheet.Range} The range.
	 * @throws {Error} If the range is not found.
	 */
	getCalendarRange(sheet)
	{
		// Dynamically locate the first row in column A that matches /^1.*lundi$/
		const lastRow = sheet.getLastRow();
		if (lastRow < 1)
		{
			throw new Error('La feuille est vide ou ne peut pas être lue.');
		}

		const colAValues = sheet.getRange(1, 1, lastRow).getDisplayValues();
		for (let i = 0; i < lastRow; i++)
		{
			if (/^1.*lundi$/.test(colAValues[i][0]))
			{
				const startRow = i + 1;

				// Range starting at Column B of that row, 20 rows by 12 columns
				return sheet.getRange(startRow, 2, 20, 12);
			}
		}

		throw new Error('Impossible de trouver le début du planning (une cellule en colonne A commençant par "1" et finissant par "lundi").');
	}

	/**
	 * Flattens the 20x12 calendar range into a 240-element array.
	 * Iterates through columns first (left to right), then rows (top to bottom).
	 * @param {GoogleAppsScript.Spreadsheet.Range} range The 20x12 calendar range.
	 * @returns {(Date|null)[]} An array containing 240 elements (Date or null).
	 */
	getPlanningDatesFromCalendarRange(range)
	{
		const values = range.getValues();
		const rows = 20;
		const cols = 12;
		const flattened = [];

		for (let c = 0; c < cols; c++)
		{
			for (let r = 0; r < rows; r++)
			{
				const val = values[r][c];
				flattened.push(val instanceof Date ? val : null);
			}
		}

		return flattened;
	}

	/**
	 * Converts the calendar dates into a map of year to 366-element tick arrays.
	 * @param {Map<number, GoogleAppsScript.Spreadsheet.Sheet>} calendarMap Map of year to sheet.
	 * @returns {Map<number, (number|null)[]>} Map of year to 366-element array of ticks.
	 */
	convertCalendarsToTicks(calendarMap)
	{
		const resultMap = new Map;

		for (const [yearKey, sheet] of calendarMap)
		{
			const yy = yearKey % 100;
			const range = this.getCalendarRange(sheet);
			const dates = this.getPlanningDatesFromCalendarRange(range);

			for (let i = 0; i < dates.length; i++)
			{
				const date = dates[i];
				if (!date)
				{
					continue;
				}

				const mm = Math.floor(i / 20) + 1;
				const row = i % 20;
				const w = Math.floor(row / 5) + 1;
				const d = (row % 5) + 1;
				const wdt = (w * 100) + (d * 10);
				const tick = (yy * 100000) + (mm * 1000) + wdt;

				const targetYear = date.getFullYear();
				const dayIndex = this._getDayOfYear(date);

				if (!resultMap.has(targetYear))
				{
					resultMap.set(targetYear, new Array(366).fill(null));
				}

				resultMap.get(targetYear)[dayIndex] = tick;
			}
		}

		return resultMap;
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