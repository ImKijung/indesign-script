var doc = app.activeDocument;
var s = doc.stories;
// var curDate  = new Date();
// var oldStamp = curDate.getTime();
var allChanges = s.everyItem().changes.everyItem().getElements();
var nChanges = allChanges.length;

for (var i = 0; i < nChanges; i++) {
	// var cType = allChanges[i].changeType;
	// if (cType == "1799644524") {
	// 	c2 = "Deleted text";
	// } else if (cType == "1799974515") {
	// 	c2 = "Added text";
	// } else if (cType == "1800236918") {
	// 	c2 = "Moved text"
	// }
	// $.writeln(c2 + " : " + allChanges[i].properties.toSource());
	
	if (allChanges[i].changeType == ChangeTypes.INSERTED_TEXT || allChanges[i].changeType == ChangeTypes.MOVED_TEXT) {
		// allChanges[i].characters.everyItem().fillColor = "C=0 M=100 Y=100 K=0";
		
		allChanges[i].texts[0].select();
		var sel = app.selection[0];
		// $.writeln("Selected Text Bounds: " + getTextBounds(sel));
		doc.activeLayer = doc.layers.itemByName("5차");
		var myFrame = sel.parentTextFrames[0].parentPage.rectangles.add();
		myFrame.geometricBounds = getTextBounds(sel);
		myFrame.strokeColor = "1차";
		myFrame.strokeWeight = "1pt"
	} else if (allChanges[i].changeType == ChangeTypes.DELETED_TEXT) {
		allChanges[i].insertionPoints[-1].contents = "†";
	}
}

app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
app.findTextPreferences.findWhat = "†";
var myDeleted = doc.findText();

for (var j=0; j<myDeleted.length; j++) {
	myDeleted[j].select();
	var selected = app.selection[0];
	doc.activeLayer = doc.layers.itemByName("5차");
	var myFrame = sel.parentTextFrames[0].parentPage.rectangles.add();
	myFrame.geometricBounds = getTextBounds(sel);
	myFrame.strokeColor = "1차";
	myFrame.strokeWeight = "1pt"
	myDeleted[j].contents = "";
}


function getTextBounds(t){
    var y1, x1, y2, x2
    y1 = t.baseline - sel.ascent;
    x1 = t.horizontalOffset
    y2 = t.baseline + sel.descent;
    x2 = t.endHorizontalOffset
    return [y1,x1,y2,x2]
}


// var cellc = doc.stories.everyItem().tables.everyItem().cells.everyItem().getElements();
// for (var n=0; n<cellc.length; n++) {
// 	var cellChng = cellc[n].changes;
// 	for (var l=0; l<cellChng.length; l++) {
// 		var cType = cellChng[l].changeType;
// 		if (cType == "1799644524") {
// 			c2 = "Deleted text";
// 		} else if (cType == "1799974515") {
// 			c2 = "Added text";
// 		} else if (cType == "1800236918") {
// 			c2 = "Moved text"
// 		}
// 		var celPara = cellChng[l].paragraphs;

// 		for (var v=0; v<celPara.length; v++) {
// 			$.writeln(l + " - " + c2 + " : " + celPara[v].contents);
// 		}
// 	}
// }