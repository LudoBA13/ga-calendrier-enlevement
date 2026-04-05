/**
 * @OnlyCurrentDoc
 */

function onOpen()
{
	const ui = SpreadsheetApp.getUi();
	ui.createMenu('Planning')
		.addItem('Créer un nouveau planning', 'createNewCalendar')
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

	// Read the 4 cells below to check their current values
	const targetRange = sheet.getRange(row + 1, col, 4, 1);
	const targetValues = targetRange.getValues();

	for (let i = 0; i < 4; i++)
	{
		const currentValue = targetValues[i][0];
		
		// If a next row is empty OR is a date, it should be modified
		if (currentValue === '' || currentValue instanceof Date)
		{
			const newValue = new Date(value);
			newValue.setDate(newValue.getDate() + (i + 1));
			targetValues[i][0] = newValue;
		}
	}

	targetRange.setValues(targetValues);
}
