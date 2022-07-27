var docs = app.documents;
var dest_file = "007_Copyright.indd";
for (var i=0; i < docs.length; i++){
	if (docs[i].name == dest_file) {
		app.activeDocument = docs[i];
	}
}
var rtlcode = String.fromCharCode(0x200F);
app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

//Set the find options.
app.findChangeTextOptions.includeFootnotes = false;
app.findChangeTextOptions.includeHiddenLayers = false;
app.findChangeTextOptions.includeLockedLayersForFind = false;
app.findChangeTextOptions.includeLockedStoriesForFind = false;
app.findChangeTextOptions.includeMasterPages = false;

app.findTextPreferences.findWhat = "Wi-Fi®";
var found01 = app.activeDocument.findText();
found01[0].insertionPoints[-1].contents= rtlcode;

app.findTextPreferences.findWhat = "Wi-Fi Direct™";
var found02 = app.activeDocument.findText();
found02[0].insertionPoints[-1].contents= rtlcode;

app.findTextPreferences.findWhat = "Wi-Fi CERTIFIED™";
var found03 = app.activeDocument.findText();
found03[0].insertionPoints[-1].contents= rtlcode;

app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

for (var i=0;i<docs.length;i++){
	var openDocs = app.documents.everyItem().getElements();
    app.activeDocument = openDocs[openDocs.length-1];
	applyLtR();
} 
alert("Left to Right 적용 완료");

function applyLtR () {
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;

	//Set the find options.
	app.findChangeGrepOptions.includeFootnotes = false;
	app.findChangeGrepOptions.includeHiddenLayers = false;
	app.findChangeGrepOptions.includeLockedLayersForFind = false;
	app.findChangeGrepOptions.includeLockedStoriesForFind = false;
	app.findChangeGrepOptions.includeMasterPages = false;

	var greps = [
		{"findWhat":"Samsung account"},
		{"findWhat":"MirrorLink"},
		{"findWhat":"Dolby Atmos"},
		{"findWhat":"Smart Lock"},
		{"findWhat":"Always On Display"},
		{"findWhat":"FaceWidgets"},
		{"findWhat":"Google Play"},
		{"findWhat":"Samsung Pass"},
		{"findWhat":"Secure Wi-Fi"},
		{"findWhat":"Galaxy Store"},
		{"findWhat":"Samsung Cloud"},
		{"findWhat":"Secure Folder"},
		{"findWhat":"Samsung Notes"},
		{"findWhat":"Smart Switch"},
		{"findWhat":"Google account"},
		{"findWhat":"Bixby Routines"},
		{"findWhat":"Game Launcher"},
		{"findWhat":"Dual Messenger"},
		{"findWhat":"Voice Assistant"},
		{"findWhat":"Bixby Vision"},
		{"findWhat":"Bixby Home"},
		{"findWhat":"Samsung Pay"},
		{"findWhat":"Samsung Health"},
		{"findWhat":"Galaxy Wearable"},
		{"findWhat":"Samsung Members"},
		{"findWhat":"screen mirroring "},
		{"findWhat":"Bluetooth®"},
		{"findWhat":"Wi-Fi®"},
		{"findWhat":"Wi-Fi Direct™"},
		{"findWhat":"Wi-Fi CERTIFIED™"},
	]
	for(var m=0; m<greps.length; m++) {
		app.findGrepPreferences.findWhat = greps[m].findWhat;
		//app.changeGrepPreferences.changeTo = greps[m].changeTo;
		var found = app.activeDocument.findGrep();
		for (var i=0;i<found.length;i++) {
			found[i].select();
			var selection = app.selection[0];
			selection.texts.everyItem().characterDirection = CharacterDirectionOptions.LEFT_TO_RIGHT_DIRECTION;
		}
	}
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
}