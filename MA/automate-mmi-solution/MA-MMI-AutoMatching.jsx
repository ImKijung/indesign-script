#targetengine "session";
#include "functions/applyID.jsx";
#include "functions/change-lang.jsx";
#include "functions/removeid.jsx";
#include "functions/eXportAddmmi.jsx";

var w = new Window ("palette", "MA MMI AutoMatching Script v.0.1.1");
var myCount = 0;
group = w.add ('group {orientation: "column"}');
panel_00 = group.add ('panel {text: ""}');
	panel_00.orientation = "row";
	var mmitxt = panel_00.add ("statictext", undefined, "MMI DB:");
	var path = panel_00.add ("statictext", [0, 0, 250, 25], "");
	var btn_00 = panel_00.add("button", [0, 0, 50, 25], "열기");
panel_01 = group.add ('panel {text: "언어 선택"}');
	panel_01.orientation = "row";
	var langDropdown = panel_01.add("dropdownlist", [0, 0, 117, 25]);
	var langlist = [];	
	var btn_01 = panel_01.add("button", [0, 0, 117, 25], "MMI 언어 변경하기");
	var btn_02 = panel_01.add("button", [0, 0, 117, 25], "MMI ID 적용하기");
panel_02 = group.add ('panel {text: ""}');
	panel_02.orientation = "row";
	var btn_03 = panel_02.add("button", [0, 0, 117, 25], "수동 ID 적용");
	var btn_04 = panel_02.add("button", [0, 0, 117, 25], "ID 모두 삭제");
	var btn_05 = panel_02.add("button", [0, 0, 117, 25], "미사용 ID 삭제");
panel_03 = group.add ('panel {text: " "}');
	panel_03.orientation = "row";
	var btn_06 = panel_03.add("button", [0, 0, 117, 25], "ID 확인하기");
	var btn_07 = panel_03.add("button", [0, 0, 117, 25], "미적용 MMI 추출");
	var btn_08 = panel_03.add("button", [0, 0, 117, 25], "표시기 변경");
panel_04 = group.add ('panel {text: "ID 찾기"}');
	panel_04.orientation = "row";
	var btn_09 = panel_04.add("button", [0, 0, 117, 25], "같은 ID 찾기");
	var btn_10 = panel_04.add("button", [0, 0, 117, 25], "-");
	var btn_11 = panel_04.add("button", [0, 0, 117, 25], "-");

btn_00.onClick = function() {
	langDropdown.removeAll();
	var mmiDB = File.openDialog("osd 용어 txt 파일을 선택하세요.", "*.txt", false);
	if (mmiDB == null) {
		exit();
	}
	path.text = decodeURI(mmiDB);
	var mmiFindChnageFile = new File(mmiDB);
	var myResult = mmiFindChnageFile.open("r", undefined, undefined);
	var myLine = mmiFindChnageFile.readln();
	var mmiFindChangeArray = myLine.split("\t");
	for (var i=1; i<mmiFindChangeArray.length; i++) {
		var mmilist = mmiFindChangeArray[i];
		//alert(mmiFindChangeArray[i]);
		var langlist = langDropdown.add("item", mmilist);
	}
	langDropdown.selection = 9;
}

btn_01.onClick = function() {
	var selectLang = langDropdown.selection.index + 1;
	var osdDB = path.text;
	

	//문서가 하나 열려있을 경우
	var docs = app.documents;
	if (docs.length == 1) {
		// $.writeln(osdDB, selectLang);
		changeLang(docs[0], osdDB, selectLang);
		alert("완료합니다.");
	}
	//문서가 여러개 열려있을 경우
	else if (docs.length > 1) {
		for (i=0; i<docs.length; i++) {
			changeLang(docs[i], osdDB, selectLang);
		}
		alert("완료합니다.");
	}
	//북파일만 열려 있을 경우
	else if (app.books.length > 0) {
		var myBook = app.activeBook;
		var mybookContents = myBook.bookContents.everyItem().getElements();
		for (var i=0; i<mybookContents.length; i++) {
			var myPath = mybookContents[i].fullName;
			var myFile = File(myPath);
			if (myFile.name.search("Cover") > 0 || myFile.name.search("TOC") > 0 || myFile.name.search("Copyright") > 0) {
				continue;
			} else {
				app.open(myFile);
			}
		}
		for (i=0; i<docs.length; i++) {
			changeLang(docs[i], osdDB, selectLang);
		}
		alert("완료합니다.");
	}
	else alert("인디자인 문서 파일 또는 북 파일을 실행하세요.");
}

btn_02.onClick = function() {
	//문서가 하나 열려있을 경우
	var docs = app.documents;
	if (docs.length == 1) {
		applyID(docs[0]);
		removeUndefinedID(docs[0]);
		alert("완료합니다.");
	}
	//문서가 여러개 열려있을 경우
	else if (docs.length > 1) {
		for (i=0; i<docs.length; i++) {
			applyID(docs[i]);
			removeUndefinedID(docs[i]);
		}
		alert("완료합니다.");
	}
	//북파일만 열려 있을 경우
	else if (app.books.length > 0) {
		var myBook = app.activeBook;
		var mybookContents = myBook.bookContents.everyItem().getElements();
		for (var i=0; i<mybookContents.length; i++) {
			var myPath = mybookContents[i].fullName;
			var myFile = File(myPath);
			if (myFile.name.search("Cover") > 0 || myFile.name.search("TOC") > 0 || myFile.name.search("Copyright") > 0) {
				continue;
			} else {
				app.open(myFile);
			}
		}
		for (i=0; i<docs.length; i++) {
			applyID(docs[i]);
			removeUndefinedID(docs[i]);
		}
		alert("완료합니다.");
	}
	else alert("인디자인 문서 파일 또는 북 파일을 실행하세요.");
}

btn_03.onClick = function() {
	var newID = prompt ("", "", "수동으로 적용할 ID를 입력하세요.");
	var doc = app.activeDocument;
	var Cond = doc.conditions;
	if (!doc.conditions.item(newID).isValid) {
		// $.writeln("add " + mmiIDx);
		Cond.add({
			name: newID,
			indicatorColor: UIColors.GRID_GREEN,
			indicatorMethod: ConditionIndicatorMethod.useHighlight
		});
		app.selection[0].appliedConditions = doc.conditions.item(newID);
	} else {
		alert("ID가 존재합니다. 기존 ID를 적용합니다.");
		app.selection[0].appliedConditions = doc.conditions.item(newID);
		doc.conditions.item(newID).indicatorColor = UIColors.GRID_GREEN;
	}
}

btn_04.onClick = function() {
	RemoveCondition();
	alert("완료합니다.");
}

btn_05.onClick = function() {
	removeUnusedConditions();
	alert("완료합니다.");
}

btn_06.onClick = function() {
	var doc = app.activeDocument;
	var mID = "";
	var mySel = app.selection[0];
	if (mySel.contents == "") {
		// alert("선택된 텍스트가 없습니다.");
		panel_03.text = "ID : None";
	}
	if (app.selection[0].appliedConditions[0] == undefined) {
		panel_03.text = "ID : None";
	} else {
		mID = mySel.appliedConditions[0].name
		panel_03.text = "ID : " + mID;
	}
}

btn_07.onClick = function() {
	if (path.text == "") {
		alert("MMI DB txt 파일을 선택하세요.");
		exit();
	}
	var osdDB = path.text;
	var txtFile = File (osdDB);
	txtFile.encoding = 'UTF-8';
	txtFile.open("r");
	var fileCotentsString = txtFile.read();
	var mmiFindChangeArray = fileCotentsString.split("\n");
	txtFile.close();
	var lineArray = mmiFindChangeArray[0].split("\t");
	// $.writeln(lineArray[0].replace("ID:",""));
	var allIDs = lineArray[0].replace("ID:","");
	allIDs *= 1;
	allIDs = allIDs + 1;
	var docs = app.documents;
	if (docs.length > 0) {
		alert ("열려있는 문서를 모두 닫은 후 북 파일만 열어놓은 상태에서 실행하세요.");
		exit();
	} else if (app.books.length > 0) {
		var myBook = app.activeBook;
		var mybookContents = myBook.bookContents.everyItem().getElements();
		for (var i=0; i<mybookContents.length; i++) {
			var myPath = mybookContents[i].fullName;
			var myFile = File(myPath);
			if (myFile.name.search("Cover") > 0 || myFile.name.search("TOC") > 0 || myFile.name.search("Copyright") > 0) {
				continue;
			} else {
				app.open(myFile);
			}
		}
		eXportAddmmi(allIDs);
		for (j=0; j<docs.length; j++) {
			docs[i].close(SaveOptions.YES);
		}
		alert("완료합니다.");
	} else alert("열려있는 북 파일이 없습니다.");
}

btn_08.onClick = function() {
	var doc = app.activeDocument;
	var stausOfindicator = doc.conditionalTextPreferences.showConditionIndicators

	if (stausOfindicator == 1698908520) {
		doc.conditionalTextPreferences.showConditionIndicators = ConditionIndicatorMode.SHOW_INDICATORS;
		panel_03.text = "표시기 : 표시";
	}
	else if (stausOfindicator == 1698908531) {
		doc.conditionalTextPreferences.showConditionIndicators = ConditionIndicatorMode.SHOW_AND_PRINT_INDICATORS;
		panel_03.text = "표시기 : 표시 및 인쇄";
	}
	else if (stausOfindicator == 1698908528) {
		doc.conditionalTextPreferences.showConditionIndicators = ConditionIndicatorMode.HIDE_INDICATORS;
		panel_03.text = "표시기 : 숨기기";
	}
}

btn_09.onClick = function () {
	var mySelID = app.selection[0].appliedConditions[0];
	if (mySelID == undefined) {
		alert("선택한 텍스트에 ID가 적용되어 있지 않습니다.");
		myCount = 0;
		panel_04.text = "None";
		exit();
	}

	//reset search
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	//set find options
	app.findChangeTextOptions.includeFootnotes = false;
	app.findChangeTextOptions.includeHiddenLayers = false;
	app.findChangeTextOptions.includeLockedLayersForFind = false;
	app.findChangeTextOptions.includeLockedStoriesForFind = false;
	app.findChangeTextOptions.includeMasterPages = false;

	app.findTextPreferences.appliedConditions = [ mySelID ];

	var myFound = app.activeDocument.findText();
	panel_04.text = myFound.length + "(" + (myCount + 1) + ") : " + mySelID.name;

	// $.writeln(myFound.length);
	if (myFound.length == 1) {
		alert("선택한 OSD와 같은 ID를 사용하는 텍스트가 없습니다.");
		myCount = 0;
		exit();
	} else {
		myFound[myCount].select();
		app.activeWindow.activePage = myFound[myCount].parentTextFrames[0].parentPage;
	}
	myCount ++;
	if (myCount >= myFound.length) {
		myCount = 0;
	}
}

w.show();

