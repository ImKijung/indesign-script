function updateAllTocIndex() {
	var myBook = app.activeBook;
	var myBookContents = myBook.bookContents.everyItem().getElements();
	// 경고창 비활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	
	for (var b=0; b<myBookContents.length; b++) {
		var myPath = myBookContents[b].fullName;
		var myDoc = File(myPath);
		app.open(myDoc);

	}
	var docs = app.documents;
	var paraName;

	if (myBook.name.search('ccNC') > 0) {
		for (var n=0; n<docs.length; n++) {
			if (docs[n].name.search('TOC') > 0) {
				app.activeDocument = docs[n];
				updateToc();
				removeTOC_ccNC()
			}
			if (docs[n].name.search('index') > 0) {
				app.activeDocument = docs[n];
				updateToc();
				generateIndex(docs[n]);
			}
		}
		for (var i = docs.length-1; i >= 0; i--) {
			docs[i].close(SaveOptions.YES);
		}
	} else {
		for (var n=0; n<docs.length; n++) {
			// var doc = docs[n];
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			var activeDoc = app.activeDocument;
			if (activeDoc.name.indexOf("Cover") != -1 || activeDoc.name.indexOf("QRG") != -1 || activeDoc.name.indexOf("Readmefirst") != -1 || activeDoc.name.indexOf("Trademarks") != -1 || activeDoc.name.indexOf("Opensource") != -1) {
				// $.writeln(activeDoc.name + " pass");
				continue;
			} else if (activeDoc.name.indexOf("TOC") != -1) { // 목차 파일인 경우
				updateToc();
				paraName = "TOC2";
				removeTOCcs(paraName);
				paraName = "TOC3";
				removeTOCcs(paraName);
			} else if (activeDoc.name.indexOf("Index") != -1) { // 인덱스 파일인 경우
				updateToc();
				generateIndex(activeDoc);
			} else { // 그 외 도비라 파일인 경우
				updateToc();
				paraName = "TOC-Chapter";
				removeTOCcs(paraName);
			}
		}
	}
	// 경고창 활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	alert("완료합니다.");
}

function updateToc() {
	var myDoc = app.activeDocument;
	var s = myDoc.stories;
	for (var i=0; i<s.length; i++) {
		try {
			if (s[i].storyType == StoryTypes.TOC_STORY) {
				s[i].textContainers[0].select();
				app.scriptMenuActions.itemByID(71442).invoke();
				// $.writeln(myDoc.name + " update TOC");
			} else continue;
		} catch(err) {
			alert(myDoc.name + " : " + err.line + " : " + err);
			exit();
		}
	}
}
function removeTOC_ccNC() {
	var doc = app.activeDocument;
	// $.writeln(doc.name);
	var removes = [ 'C_Below_Heading', 'C_Below_Chapter' ];
	
	for (var j=0;j<removes.length;j++) {
		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
		app.findGrepPreferences.findWhat = '(.+)';
		app.findGrepPreferences.appliedCharacterStyle = doc.characterStyles.itemByName(removes[j])
		
		var myFound = doc.findGrep();
		for (var k=0;k<myFound.length;k++) {
			myFound[k].contents = '';
		}
		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	}
	doc.save();
}
function removeTOCcs(paraName) {
	var myDoc = app.activeDocument;
	var ps = myDoc.allParagraphStyles;
	for (var i=0; i<ps.length; i++) {
		if (ps[i].name == paraName) {
			var tocPS = ps[i];
		}
	}
	var removeCS = myDoc.characterStyles.itemByName('C_Below_Heading');
	var None = myDoc.characterStyles[0];

	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	try {
		if (paraName == "TOC2" || paraName == "TOC3") {
			app.findGrepPreferences.findWhat = "(.+?)(\\t\\d{1,}\\-\\d{1,})";
		} else {
			app.findGrepPreferences.findWhat = "(.+?)$";
		}
		app.findGrepPreferences.appliedParagraphStyle = tocPS;
		app.findGrepPreferences.appliedCharacterStyle = removeCS;
		if (paraName == "TOC2" || paraName == "TOC3") {
			app.changeGrepPreferences.changeTo = "$2";
		} else {
			app.changeGrepPreferences.changeTo = "";
		}
		if (paraName == "TOC2" || paraName == "TOC3") {
			app.changeGrepPreferences.appliedCharacterStyle = None;
		}
		myDoc.changeGrep();
	} catch(err) {
		alert(myDoc.name + " : " + err.line + " : " + err);
		exit();
	}
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}

function generateIndex(doc) {
	var myIndex = doc.indexes[0];
	doc.indexGenerationOptions.includeBookDocuments = true;
	doc.indexGenerationOptions.title = "";
	myIndex.generate();

	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	app.findGrepPreferences.findWhat = "로마자\r";
	app.changeGrepPreferences.changeTo = "";
	doc.changeGrep();
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;

	// $.writeln(doc.name + " index update")
}
