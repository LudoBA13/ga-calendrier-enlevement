/**
 * @OnlyCurrentDoc
 */

function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Planning')
		.addItem('Créer un nouveau planning', 'createNewCalendar')
		.addItem('Mettre en cache', 'cacheCurrentPlanning')
		.addToUi();
}

function createNewCalendar()
{
	const ui = SpreadsheetApp.getUi();
	const response = ui.prompt('Nouveau planning', 'Entrez l\'année :', ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() !== ui.Button.OK)
	{
		return;
	}

	const year = response.getResponseText().trim();
	if (!year || isNaN(year))
	{
		ui.alert('Erreur', 'Veuillez entrer une année valide.', ui.ButtonSet.OK);
		return;
	}

	const sheetName = 'Calendrier' + year;
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	let sheet = ss.getSheetByName(sheetName);

	if (sheet)
	{
		ui.alert('Information', 'Le planning ' + year + ' existe déjà.', ui.ButtonSet.OK);
		ss.setActiveSheet(sheet);
		return;
	}

	const template = ss.getSheetByName('CalendrierModèle');
	if (!template)
	{
		ui.alert('Erreur', 'La feuille \'CalendrierModèle\' est introuvable.', ui.ButtonSet.OK);
		return;
	}

	sheet = template.copyTo(ss);
	sheet.setName(sheetName);
	sheet.getRange('A1').setValue(year);
	ss.setActiveSheet(sheet);
	ui.alert('Succès', 'Le planning ' + year + ' a été créé.', ui.ButtonSet.OK);
}

/**
 * Caches the planning data from the current active sheet.
 */
function cacheCurrentPlanning()
{
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();
	const sheet = ss.getActiveSheet();

	const result = performCaching(sheet);

	if (result.success)
	{
		ui.alert('Succès', result.message, ui.ButtonSet.OK);
	}
	else
	{
		ui.alert('Erreur', result.message, ui.ButtonSet.OK);
	}
}

/**
 * Logic to perform caching for a specific sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {{success: boolean, message: string}}
 */
function performCaching(sheet)
{
	const sheetName = sheet.getName();
	const match = sheetName.match(/^Calendrier(20\d+)$/);

	if (!match)
	{
		return { success: false, message: 'Cette action ne peut être effectuée que sur une feuille de planning (ex: Calendrier2026).' };
	}

	const year = parseInt(match[1]);
	const range = getCalendarRange(sheet);

	if (!range)
	{
		return { success: false, message: 'Impossible de trouver le début du planning (une cellule en colonne A commençant par "1" et finissant par "lundi").' };
	}

	try
	{
		savePlanning(year, range);
		return { success: true, message: 'Le planning ' + year + ' a été mis en cache.' };
	}
	catch (error)
	{
		return { success: false, message: 'Une erreur est survenue lors de la mise en cache : ' + error.message };
	}
}

/**
 * Finds the 20x12 calendar range in the given sheet.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The sheet to search in.
 * @returns {GoogleAppsScript.Spreadsheet.Range|null} The range or null if not found.
 */
function getCalendarRange(sheet)
{
	// Dynamically locate the first row in column A that matches /^1.*lundi$/
	const lastRow = sheet.getLastRow();
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
		return null;
	}

	// Range starting at Column B of that row, 20 rows by 12 columns
	return sheet.getRange(startRow, 2, 20, 12);
}

/**
 * Triggered when a cell is modified.
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e)
{
	const range = e.range;
	const sheet = range.getSheet();
	const sheetName = sheet.getName();

	// Check sheet name pattern
	if (!/^Calendrier20\d+$/.test(sheetName))
	{
		return;
	}

	const row = range.getRow();
	const col = range.getColumn();
	const cellA = sheet.getRange(row, 1);
	const cellAValue = cellA.getDisplayValue();

	// Read the cell in the A column and check if it ends with "lundi"
	if (cellAValue.endsWith('lundi'))
	{
		onEditMonday(range, sheet, row, col);
	}

	// Automatically cache if the edited value is a Date
	if (range.getValue() instanceof Date)
	{
		performCaching(sheet);
	}
}

/**
 * Handles the logic when a cell on a "lundi" row is edited.
 * @param {GoogleAppsScript.Spreadsheet.Range} range The modified range.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet The active sheet.
 * @param {number} row The row number of the edit.
 * @param {number} col The column number of the edit.
 */
function onEditMonday(range, sheet, row, col)
{
	const value = range.getValue();

	// Only modify if the new value of the edited cell is a date
	if (!(value instanceof Date))
	{
		return;
	}

	const cellAValue = sheet.getRange(row, 1).getDisplayValue();
	const match = cellAValue.match(/^([1234])/);
	if (!match)
	{
		return;
	}

	const n = parseInt(match[1]);
	const count = ((5 - n) * 5) - 1;

	// Read the cells below to check their current values
	const targetRange = sheet.getRange(row + 1, col, count, 1);
	const targetValues = targetRange.getValues();

	let daysToAdd = 0;
	for (let i = 1; i <= count; i++)
	{
		// After 4 consecutive days (Tue, Wed, Thu, Fri), skip to next Monday
		if (i % 5 === 0)
		{
			daysToAdd += 3;
		}
		else
		{
			daysToAdd += 1;
		}

		const currentIndex = i - 1;
		const currentValue = targetValues[currentIndex][0];

		// If a next row is empty OR is a date, it should be modified
		if (currentValue === '' || currentValue instanceof Date)
		{
			const newValue = new Date(value);
			newValue.setDate(newValue.getDate() + daysToAdd);
			targetValues[currentIndex][0] = newValue;
		}
	}

	targetRange.setValues(targetValues);
}
