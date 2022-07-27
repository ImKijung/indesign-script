var doc = app.activeDocument;

app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

app.findGrepPreferences.findWhat = "[^\r]+";  

var myFound = doc.findGrep();

for (i=0; i<myFound.length; i++) {
	$.writeln(myFound[i].contents);
}

app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;