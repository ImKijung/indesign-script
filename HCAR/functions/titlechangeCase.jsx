function titleChangeCaseUpper() {
	var myDoc = app.activeDocument;
	// if equipped, for europe, expect eroupe 
	var excludeWords = [
		"at", "in", "on", "upon", "for", "by", "of", "with", "during", "before", "after", "following", "until", "since", "within", "over", "behind", "ahead of", "because of", "despite", "about", "from", "to", "into", "through", "throughout", "toward", "towards", "among", "between", "across", "around", "opposite", "past", "near", "along", "alongside", "except", "except for", "besides", "instead of", "without", "as", "under", "beneath", "beyond", "above", "below", "out of", "without", "a", "an", "the", "(if", "equipped)", "(for", "europe)", "(expect"
	]

	var excludes = excludeWords.join("|");
	var headingStyles = /^Heading1|^Heading1.+?|^Heading2|^Heading2.+?|^Heading3|^Heading3.+?|^Chapter$|^Chapter_Con$/
	var ps = myDoc.allParagraphStyles;
	var paraStyleName
	for (var i=0; i<ps.length; i++) {
		if (ps[i].name.match(headingStyles)) {
			// $.writeln(ps[i].name);
			paraStyleName = ps[i].name;
			upperCaseTitle(myDoc, paraStyleName, excludes);
		};
	}
}

function upperCaseTitle(myDoc, paraStyleName, excludes) {
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	try {
		// app.findGrepPreferences.findWhat = "^(.+?)$";
		app.findGrepPreferences.appliedParagraphStyle = myDoc.paragraphStyles.itemByName(paraStyleName);
		var myFound = myDoc.findGrep();
		for (var n=0; n<myFound.length; n++) {
			// $.writeln(paraStyleName + " : " + myFound[n].contents);
			var Words = myFound[n].words;
			for (var j=0;j<Words.length;j++) {
				if (Words[j].contents == "Ev") {
					Words[j].contents = "EV";
				}
				if (excludes.indexOf(Words[j].contents) != -1) {
					continue;
				} else {
					Words[j].characters[0].changecase(ChangecaseMode.uppercase);
				}
			}
		}
	} catch(err) {
		alert(myDoc.name + " : " + err.line + " : " + err);
		exit();
	}
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}