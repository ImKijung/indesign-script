var doc = app.activeDocument;

var finds = "(?!\\r)^~a\\s(\\w|\\()";
var finds1 = "\\.\\s~a\\s(\\w|\\()";
var nwb = String.fromCharCode(0xFEFF);


app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

//Set the find options.
app.findChangeGrepOptions.includeFootnotes = false;
app.findChangeGrepOptions.includeHiddenLayers = false;
app.findChangeGrepOptions.includeLockedLayersForFind = false;
app.findChangeGrepOptions.includeLockedStoriesForFind = false;
app.findChangeGrepOptions.includeMasterPages = false;

app.findGrepPreferences.findWhat = finds;

var found = doc.findGrep();
var count = 0, count1 = 0;

for (var i=found.length-1; i>=0; i--) {
	if (found[i].paragraphs[0].characters[0].contents == nwb) {
		// $.writeln(i + ":: Error 111");
	} else {
		found[i].insertionPoints[0].contents = nwb;
		count++;
	}
}

app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

app.findGrepPreferences.findWhat = finds1;

var found1 = doc.findGrep();

for (var j=found1.length-1; j>=0; j--) {
	if (found1[j].characters[2].contents == nwb) {
		// $.writeln(j + ":: Error 2222");
	} else {
		found1[j].insertionPoints[2].contents = nwb;
		count1++;
	}
}

// $.writeln("중간-" + count1 + " ::: " + "else " + count);

app.findGrepPreferences = NothingEnum.nothing;
app.changeGrepPreferences = NothingEnum.nothing;

alert("완료합니다.");