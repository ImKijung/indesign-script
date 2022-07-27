var doc = app.activeDocument;
var osd1= doc.characterStyles.item("C_OSD");
var osd2= doc.characterStyles.item("C_OSD-NoBold");
var osdlist = [ osd1, osd2 ]

//reset search
app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

//set find options
app.findChangeTextOptions.includeFootnotes = false;
app.findChangeTextOptions.includeHiddenLayers = false;
app.findChangeTextOptions.includeLockedLayersForFind = false;
app.findChangeTextOptions.includeLockedStoriesForFind = false;
app.findChangeTextOptions.includeMasterPages = false;

//find OSD style and apply ID

for (var i=0;i<osdlist.length;i++) {
	app.findTextPreferences.appliedCharacterStyle = osdlist[i];
	var found = app.findText();
	$.writeln(found.length);
	
}