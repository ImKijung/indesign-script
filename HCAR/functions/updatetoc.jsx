updateIndex();
function updateIndex() {
	var myBook = app.activeBook;
	var myBookContents = myBook.bookContents.everyItem().getElements();
	// 경고창 비활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	for (var b=0; b<myBookContents.length; b++) {
		if (myBookContents[b].name.indexOf("INDEX") != -1) {
			var myPath = myBookContents[b].fullName;
			var myDoc = File(myPath);
			app.open(myDoc);
		}
	}
	
	updateToc();
	var paraNames = [ "TOC_Index1", "TOC_Index2" ];
	for (var x=0; x<paraNames.length; x++) {
		removeCarrigereturn(paraNames[x]);
		removeTOC1(paraNames[x]);
		removeTOC2(paraNames[x]);
		RemoveforceLinebreak(paraNames[x]);
	}
	// generateIndex(activeDoc);
	generateIndex();
	app.activeDocument.save();
	// 경고창 활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	alert("완료합니다.");
}

function updateAllToc(opt) {
	// 경고창 비활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	var docs = app.documents;
	var x, paraNames;
	var completeDoc = [];

	if (opt == 0) {
		var docLength = 1;
	} else if (opt == 1) {
		var docLength = docs.length;
	}

	for (var n=0; n<docLength; n++) {
		// var doc = docs[n];
		var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, docs.length);
		myProg.pbar.preferredSize.width = 300;
		myProg.show();
		if (docLength == 1 ) {
			var activeDoc = app.activeDocument;
			completeDoc.push(activeDoc.name);
		} else if (docLength > 1) {
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			var activeDoc = app.activeDocument;
			completeDoc.push(activeDoc.name);
		}

		if (activeDoc.name.indexOf("INDEX") != -1) { // 인덱스 파일인 경우
			// 추가 코드
			// continue;
			myProg.text = activeDoc.name + ' 업데이트 중...';
			myProg.pbar.value = n + 1;
			
			paraNames = [ "TOC_Index1", "TOC_Index2" ];
			updateToc();
			for (x=0; x<paraNames.length; x++) {
				removeCarrigereturn(paraNames[x]);
				removeTOC1(paraNames[x]);
				removeTOC2(paraNames[x]);
				RemoveforceLinebreak(paraNames[x]);
			}
			// generateIndex(activeDoc);
			generateIndex();
			activeDoc.save();
		} else { // 그 외 
			// 추가 코드
			myProg.text = activeDoc.name + ' 업데이트 중...';
			myProg.pbar.value = n + 1;
			updateToc();
			paraNames = ["TOC1", "TOC2", "TOC3"];
			for (x=0; x<paraNames.length; x++) {
				removeCarrigereturn(paraNames[x]);
				removeTOC1(paraNames[x]);
				removeTOC2(paraNames[x]);
				RemoveforceLinebreak(paraNames[x]);
			}
			activeDoc.save();
		}
	}
	// 종료
	myProg.close();
	// 경고창 활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	// $.writeln(completeDoc);
	alert("작업 완료한 문서 : " + completeDoc);
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

function removeCarrigereturn(paraName) {
	var myDoc = app.activeDocument;
	var ps = myDoc.allParagraphStyles;
	for (var i=0; i<ps.length; i++) {
		if (ps[i].name == paraName) {
			var tocPS = ps[i];
		}
	}
	var removeCS = myDoc.characterStyles.itemByName('C_TOC_Del');
	if (!removeCS.isValid) {
		alert("C_TOC_Del 문자 스타일이 없습니다. 문자 스타일을 추가한 다음 다시 실행하세요.");
		exit();
	}
	var None = myDoc.characterStyles[0];

	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	try {
		app.findGrepPreferences.findWhat = "(\\r)";
		app.findGrepPreferences.appliedParagraphStyle = tocPS;
		app.findGrepPreferences.appliedCharacterStyle = removeCS;
		app.changeGrepPreferences.changeTo = "$1";
		app.changeGrepPreferences.appliedCharacterStyle = None;
		myDoc.changeGrep();
	} catch(err) {
		alert(myDoc.name + " : " + err.line + " : " + err);
		exit();
	}
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}

function removeTOC1(paraName) {
	var myDoc = app.activeDocument;
	var ps = myDoc.allParagraphStyles;
	for (var i=0; i<ps.length; i++) {
		if (ps[i].name == paraName) {
			var tocPS = ps[i];
		}
	}
	var removeCS = myDoc.characterStyles.itemByName('C_TOC_Del');
	if (!removeCS.isValid) {
		alert("C_TOC_Del 문자 스타일이 없습니다. 문자 스타일을 추가한 다음 다시 실행하세요.");
		exit();
	}
	var None = myDoc.characterStyles[0];

	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	try {
		app.findGrepPreferences.findWhat = "(\\([^>]+\\))(\\t\\d{1,}\\-\\d{1,})";
		app.findGrepPreferences.appliedParagraphStyle = tocPS;
		app.findGrepPreferences.appliedCharacterStyle = removeCS;
		app.changeGrepPreferences.changeTo = "$2";
		app.changeGrepPreferences.appliedCharacterStyle = None;
		myDoc.changeGrep();
	} catch(err) {
		alert(myDoc.name + " : " + err.line + " : " + err);
		exit();
	}
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}

function removeTOC2(paraName) {
	var myDoc = app.activeDocument;
	var ps = myDoc.allParagraphStyles;
	for (var i=0; i<ps.length; i++) {
		if (ps[i].name == paraName) {
			var tocPS = ps[i];
		}
	}
	var removeCS = myDoc.characterStyles.itemByName('C_TOC_Del');
	var None = myDoc.characterStyles[0];

	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	try {
		app.findTextPreferences.findWhat = "";
		app.findTextPreferences.appliedParagraphStyle = tocPS;
		app.findTextPreferences.appliedCharacterStyle = removeCS;
		app.changeTextPreferences.changeTo = "";
		// app.changeTextPreferences.appliedCharacterStyle = None;
		myDoc.changeText();
	} catch(err) {
		alert(myDoc.name + " : " + err.line + " : " + err);
		exit();
	}
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
}

function RemoveforceLinebreak(paraName) {
	var myDoc = app.activeDocument;
	var ps = myDoc.allParagraphStyles;
	for (var i=0; i<ps.length; i++) {
		if (ps[i].name == paraName) {
			var tocPS = ps[i];
		}
	}
	
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	try {
		app.findGrepPreferences.findWhat = "(\\s\\n)";
		app.findGrepPreferences.appliedParagraphStyle = tocPS;
		app.changeGrepPreferences.changeTo = "";
		myDoc.changeGrep();
	} catch(err) {
		alert(myDoc.name + " : " + err.line + " : " + err);
		exit();
	}
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}

// function generateIndex(myDoc) {
// 	var myDoc = app.activeDocument;
// 	var ps = myDoc.allParagraphStyles;
// 	var paraName = "TOC_Index1";
// 	for (var i=0; i<ps.length; i++) {
// 		if (ps[i].name == paraName) {
// 			var tocPS = ps[i];
// 		}
// 	}
// 	// 색인 값 배열로 만들기
// 	var indexList = generateIndexArray(myDoc, tocPS);

// 	for (var y=0; y<indexList.length; y++) {
// 		indexAlpha(myDoc, indexList[y], tocPS);
// 	}

// 	function generateIndexArray(myDoc, tocPS) { // 색인 값 배열로 만들기
// 		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
// 		try {
// 			app.findGrepPreferences.findWhat = "^(.+?)$";
// 			app.findGrepPreferences.appliedParagraphStyle = tocPS;
// 			var myFound = myDoc.findGrep();
// 			var ToClist = [];
// 			for (var j=0; j<myFound.length; j++) {
// 				ToClist.push(myFound[j].paragraphs[0].characters[0].contents);
// 			}
// 			// $.writeln(duplicateRemove(ToClist));
// 			var indexArray = duplicateRemove(ToClist);
// 			return indexArray;

// 		} catch(err) {
// 			alert(myDoc.name + " : " + err.line + " : " + err);
// 			exit();
// 		}
// 		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
// 	}

// 	function duplicateRemove(alphabet) { // 중복 제거
// 		var m = {};
// 		var newarr = [];
// 		for (var x=0; x<alphabet.length; x++) {
// 			var v = alphabet[x];
// 			if (!m[v]) {
// 				newarr.push(v);
// 				m[v] = true;
// 			}
// 		}
// 		return newarr
// 	}

// 	function indexAlpha(myDoc, alpha, tocPS) {
// 		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
// 		try {
// 			app.findGrepPreferences.findWhat = "(^" + alpha + "(.+?)\\r)";
// 			app.findGrepPreferences.appliedParagraphStyle = tocPS;
// 			var myFound = myDoc.findGrep();
	
// 			var indexTop = myFound[0].insertionPoints[0];
// 			indexTop.contents = alpha + "\r";
// 			indexTop.appliedParagraphStyle = myDoc.paragraphStyles.itemByName("ABC_Index");
			
// 		} catch(err) {
// 			alert(myDoc.name + " : " + err.line + " : " + err);
// 			exit();
// 		}
// 		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
// 	}
// }

function generateIndex() {
	var myDoc = app.activeDocument;

	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	app.findGrepPreferences.findWhat = "(^.*?$)";
	app.findGrepPreferences.appliedParagraphStyle = myDoc.paragraphStyles.item("TOC_Index1");
	var myFound = myDoc.findGrep();
	var tempFrame;

	// revers 방식으로 진행한다.
	try {
		for (var j=myFound.length - 1; j>=0; j--) {
			var curChar = myFound[j].paragraphs[0].characters[0].contents;
			if (j == 0) {
				tempFrame = myDoc.textFrames.add();
				tempFrame.parentStory.contents = curChar + "\r";
				tempFrame.parentStory.appliedParagraphStyle = myDoc.paragraphStyles.item("ABC_Index");
				tempFrame.parentStory.texts[0].move(LocationOptions.AT_BEGINNING, myFound[j].insertionPoints[0]);
				tempFrame.remove();
			} else {
				var nextChar = myFound[j-1].paragraphs[0].characters[0].contents;
				if (curChar == nextChar) {
					continue
				} else {
					$.writeln(curChar + " - " + nextChar);
					tempFrame = myDoc.textFrames.add();
					tempFrame.parentStory.contents = curChar + "\r";
					tempFrame.parentStory.appliedParagraphStyle = myDoc.paragraphStyles.item("ABC_Index");
					tempFrame.parentStory.texts[0].move(LocationOptions.AT_BEGINNING, myFound[j].insertionPoints[0]);
					tempFrame.remove();
				}
			}
		}
	} catch(e) {
		alert(j + " - " + e + ":" + e.line);
	}
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}
