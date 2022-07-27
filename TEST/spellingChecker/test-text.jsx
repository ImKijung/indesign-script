var doc = app.activeDocument;

app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;
app.findChangeGrepOptions.properties = {
	includeFootnotes:false,
	includeHiddenLayers:false,
	includeLockedLayersForFind:false,
	includeLockedStoriesForFind:false,
	includeMasterPages:false,
}

app.findGrepPreferences.findWhat = "[^\r]+";  

var myFound = app.findGrep();
var allTxt = "";

for (var i=0; i<myFound.length; i++){
	allTxt = allTxt.concat(myFound[i].contents + "\r");
}


$.writeln(allTxt);