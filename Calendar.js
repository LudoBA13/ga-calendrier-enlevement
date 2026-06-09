/**
 * @OnlyCurrentDoc
 */

function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Calendrier')
		.addItem('Créer un nouveau calendrier', 'createNewCalendar')
		.addItem('Rafraîchir tout le stockage', 'refreshAllStorage')
		.addToUi();
}

/**
 * Refreshes all calendar storage including planning codes and ticks.
 */
function refreshAllStorage()
{
	const ui = SpreadsheetApp.getUi();
	const ss = SpreadsheetApp.getActiveSpreadsheet();

	try
	{
		const manager = new CalendarManager(ss);
		const storage = new CalendarStorage(ss);

		storage.refresh(manager);

		ui.alert('Succès', 'Le stockage a été rafraîchi avec succès (Codes et Ticks).', ui.ButtonSet.OK);
	}
	catch (error)
	{
		ui.alert('Erreur', 'Une erreur est survenue lors du rafraîchissement du stockage : ' + error.message, ui.ButtonSet.OK);
	}
}

function createNewCalendar()
{
	const ui = SpreadsheetApp.getUi();
	const response = ui.prompt('Nouveau calendrier', 'Entrez l\'année :', ui.ButtonSet.OK_CANCEL);

	if (response.getSelectedButton() !== ui.Button.OK)
	{
		return;
	}

	const year = parseInt(response.getResponseText().trim());
	if (isNaN(year) || year < 2023)
	{
		ui.alert('Erreur', 'Veuillez entrer une année valide (>= 2023).', ui.ButtonSet.OK);
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

	// Add year to 'PlanningToDate' sheet
	const planningSheet = ss.getSheetByName('PlanningToDate');
	if (planningSheet)
	{
		try
		{
			const row = year - 2020;
			const manager = new CalendarManager(ss);
			const range = manager.getCalendarRange(sheet);
			planningSheet.getRange(row, 1).setValue(year);
			planningSheet.getRange(row, 2).setFormula('=TOROW(\'' + sheetName + '\'!' + range.getA1Notation() + '; 0; 1)');
		}
		catch (error)
		{
			ui.alert('Erreur', 'Impossible d\'ajouter l\'année à la feuille \'PlanningToDate\' : ' + error.message, ui.ButtonSet.OK);
		}
	}

	ss.setActiveSheet(sheet);
	ui.alert('Succès', 'Le planning ' + year + ' a été créé.', ui.ButtonSet.OK);
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
			targetValues[currentIndex][0] = isHoliday(newValue) ? '' : newValue;
		}
	}

	targetRange.setValues(targetValues);
}

/**
 * Checks if a date is a holiday (e.g., May 1st).
 * @param {Date} date The date to check.
 * @returns {boolean} True if the date is a holiday.
 */
function isHoliday(date)
{
	return date.getMonth() === 4 && date.getDate() === 1;
}

