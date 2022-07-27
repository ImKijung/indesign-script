// var allChanges = app.documents[0].stories.everyItem().changes.everyItem().getElements();
// var allChangesCell = app.documents[0].stories.everyItem().tables.everyItem().cells.everyItem().changes.everyItem().getElements();

// $.writeln(allChanges.length + " " + allChangesCell.length);

var allChanges =app.documents[0].stories.everyItem().changes.everyItem().getElements();

var nChanges = allChanges.length;

for (var i = 0; i < nChanges; i++) {

if (allChanges[i].changeType == ChangeTypes.INSERTED_TEXT) { 
	// allChanges[i].characters.everyItem().fillColor = "C=0 M=0 Y=100 K=0";
	var ddd = allChanges[i].characters.everyItem();
	ddd.select();
	exit();
}

else if(allChanges[i].changeType == ChangeTypes.DELETED_TEXT) { }

}