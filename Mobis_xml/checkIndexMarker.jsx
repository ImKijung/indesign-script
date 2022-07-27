var myDoc = app.activeDocument;

app.findTextPreferences = NothingEnum.nothing;  
app.changeTextPreferences = NothingEnum.nothing;
app.findTextPreferences.findWhat = "<FEFF>";

var myFound = myDoc.findText();

try {
	for (var n=0; n<myFound.length; n++) {
		myFound[n].select();
		// $.writeln(findTopic(app.selection[0]).name + " or " + findReference(app.selection[0]).name);
		if ( findTopic(app.selection[0]).name == undefined && findReference(app.selection[0]).name == undefined ) {
			if (checkNotes(myDoc, app.selection[0])) {
				continue
			} else app.selection[0].remove();
		}
	}
	alert("Complete");
} catch(e) {
	// var selection = app.selection[0].insertionPoints[0].index;
	app.activeWindow.zoom = ZoomOptions.fitPage;
	alert(e.line + " : " + e + "///" + app.selection[0].insertionPoints[0].index);
}

function checkNotes (doc, marker) {
	var sIndex = marker.insertionPoints[0].index;
	var Stories = doc.stories;

	for (var i=0;i<Stories.length;i++) {
		var story = Stories[i];
		var Notes = story.notes;
		for (var j=0;j<Notes.length;j++) {
			var nIndex = Notes[j].storyOffset.insertionPoints[0].index;
			if (sIndex == nIndex) {
				return true;
			}
		}
	}
	return false;
}

function findTopic (marker) {
    var topics = app.activeDocument.indexes[0].allTopics;
    var mStory = marker.parentStory;
    var mIndex = marker.insertionPoints[0].index;
    var pRefs;
    for (var i = topics.length-1; i >= 0; i--) {
        pRefs = topics[i].pageReferences.everyItem().getElements();
        for (var j = pRefs.length-1; j >= 0; j--){
            if (pRefs[j].sourceText.parentStory == mStory && pRefs[j].sourceText.index == mIndex) {
                return pRefs[j].parent;
            }
        }
    }
    return "";
}

function findReference(marker) {
	// return marker.parent.parent.hyperlinks.constructor.name
	var myRefs = app.activeDocument.hyperlinkTextDestinations;
	var myStory = marker.parentStory;
    var myIndex = marker.insertionPoints[0].index;

	for (var k=0; k<myRefs.length; k++) {
		if (myRefs[k].destinationText.parentStory == myStory && myRefs[k].destinationText.index == myIndex) {
			return myRefs[k];
		}
	}
	return "";
}