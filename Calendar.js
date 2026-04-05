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
	
	const model = ss.getSheetByName('CalendrierModèle');
	if (!model)
	{
		ui.alert('Erreur', 'La feuille \'CalendrierModèle\' est introuvable.', ui.ButtonSet.OK);
		return;
	}
	
	sheet = model.copyTo(ss);
	sheet.setName(sheetName);
	sheet.getRange('A1').setValue(year);
	ss.setActiveSheet(sheet);
	
	ui.alert('Succès', 'Le planning ' + year + ' a été créé.', ui.ButtonSet.OK);
}
