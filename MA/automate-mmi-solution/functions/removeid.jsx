function removeUnusedConditions() {
	var cdns = app.activeDocument.conditions;
	var myWindow = new Window ('palette');
				myWindow.pbar = myWindow.add ('progressbar', undefined, 0, cdns.length);
				myWindow.pbar.preferredSize.width = 300;
				myWindow.show();
	for (i=0; i<cdns.length; i++) {
		myWindow.pbar.value = i+1;
		app.findGrepPreferences = app.changeGrepPreferences = null;
		app.findGrepPreferences.appliedConditions = [ cdns[i].name ];
		var myFound = app.activeDocument.findGrep();
		if (myFound.length == 0) cdns[i].remove();
		app.findGrepPreferences = app.changeGrepPreferences = null;
	}
}

function RemoveCondition() {
	app.activeDocument.conditions.everyItem().remove();
}

function removeUndefinedID(doc) {
	// var doc = app.activeDocument; 
	var unassign = doc.conditions.item("mmi-unassign");
	app.findTextPreferences = app.changeTextPreferences = null;
	app.findTextPreferences.appliedConditions = [ unassign ];
	var myFound = doc.findText();
	if (myFound.length == 0) unassign.remove();
	app.findTextPreferences = app.changeTextPreferences = null;
}