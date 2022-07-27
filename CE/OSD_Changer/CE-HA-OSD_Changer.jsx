#targetengine "session";
#include "functions/apply-id.jsx";
#include "functions/change-lang.jsx";
#include "functions/removeid.jsx";
#include "functions/eXportosd.jsx";

var state = 0;

var w = new Window ("palette", "CE | 생활가전 OSD Changer v.0.0.6");
	group = w.add ('group {orientation: "column"}');
	panel_00 = group.add ('panel {text: "언어 선택"}');
	panel_00.orientation = "row";
	btn_00 = panel_00.add("button", [0, 0, 100, 30], "DB txt 불러오기");
	var langDropdown = panel_00.add("dropdownlist", [0, 0, 100, 30]);
	var langlist = []
	
	btn_01 = panel_00.add("button", [0, 0, 100, 30], "ID 적용하기");
	btn_02 = panel_00.add("button", [0, 0, 100, 30], "언어 변경하기");

	panel_01 = group.add ('panel {text: "ID : "}');
	panel_01.orientation = "row";
	btn_09 = panel_01.add("button", [0, 0, 100, 30], "ID 확인하기");
	btn_11 = panel_01.add("button", [0, 0, 100, 30], "같은 ID 찾기");
	btn_08 = panel_01.add("button", [0, 0, 100, 30], "수동 ID 적용하기");
	btn_10 = panel_01.add("button", [0, 0, 100, 30], "Unused ID 삭제");
	
	panel_02 = group.add ('panel {text: "옵션"}');
	panel_02.orientation = "row";
	btn_04 = panel_02.add("button", [0, 0, 100, 30], "수동 ID 색상 변경");
	btn_03 = panel_02.add("button", [0, 0, 100, 30], "모든 ID 삭제");
	btn_12 = panel_02.add("button", [0, 0, 100, 30], "적용된 ID 추출");
	btn_05 = panel_02.add("button", [0, 0, 100, 30], "표시기");

w.show();
btn_00.onClick = function() {
	langDropdown.removeAll();

	var osdDB = File.openDialog("osd 용어 txt 파일을 선택하세요.", "*.txt", false);
	if (osdDB == null) { // osdDB 가 없을 경우
		panel_00.text = "None";
		exit();
	}
	// osdDB 가 있을 경우
	var mmiFindChnageFile = new File(osdDB);
	var myResult = mmiFindChnageFile.open("r", undefined, undefined);
	var myLine = mmiFindChnageFile.readln();
	var mmiFindChangeArray = myLine.split("\t");
	panel_00.text = decodeURI(osdDB);
	for (var i=2; i<mmiFindChangeArray.length; i++) {
		var mmilist = mmiFindChangeArray[i];
		//alert(mmiFindChangeArray[i]);
		var langlist = langDropdown.add("item", mmilist);
	}
	langDropdown.selection = 0;
}

btn_01.onClick = function() {
	var selectLang = langDropdown.selection.index + 2;
	osdDB = panel_00.text;
	w.close();
	applyID(osdDB, selectLang);
	w.show();
}

btn_02.onClick = function() {
	var selectLang = langDropdown.selection.index + 2;
	osdDB = panel_00.text
	w.close();
	changeLang(osdDB, selectLang)
	w.show();
}

btn_03.onClick = function() {
	app.activeDocument.conditions.everyItem().remove();
	alert("OSD ID를 해제했습니다.");
}

btn_04.onClick = function() {
	var doc = app.activeDocument;
	var Cond = doc.conditions;
	for (var i=0;i<Cond.length;i++) {
		if (Cond[i].indicatorColor !== UIColors.YELLOW) {
			Cond[i].indicatorColor = UIColors.GREEN;
		}
	}
}

btn_05.onClick = function() {
	var doc = app.activeDocument;
	var stausOfindicator = doc.conditionalTextPreferences.showConditionIndicators
	if (stausOfindicator == 1698908520) {
		doc.conditionalTextPreferences.showConditionIndicators = ConditionIndicatorMode.SHOW_INDICATORS;
		btn_05.text = "표시";
	}
	else if (stausOfindicator == 1698908531) {
		doc.conditionalTextPreferences.showConditionIndicators = ConditionIndicatorMode.SHOW_AND_PRINT_INDICATORS;
		btn_05.text = "표시 및 인쇄";
	}
	else if (stausOfindicator == 1698908528) {
		doc.conditionalTextPreferences.showConditionIndicators = ConditionIndicatorMode.HIDE_INDICATORS;
		btn_05.text = "숨기기";
	}
}

btn_08.onClick = function() {
	var newID = prompt ("", "", "수동으로 적용할 ID를 입력하세요.");
	var doc = app.activeDocument;
	var Cond = doc.conditions;
	if (!doc.conditions.item(newID).isValid) {
		// $.writeln("add " + mmiIDx);
		Cond.add({
			name: newID,
			indicatorColor: UIColors.GREEN,
			indicatorMethod: ConditionIndicatorMethod.useHighlight
		});
		app.selection[0].appliedConditions = doc.conditions.item(newID);
	} else {
		alert("ID가 존재합니다. 기존 ID를 적용합니다.");
		app.selection[0].appliedConditions = doc.conditions.item(newID);
		doc.conditions.item(newID).indicatorColor = UIColors.GREEN;
	}
}

btn_09.onClick = function () {
	var doc = app.activeDocument;
	var mID = "";
	var mySel = app.selection[0];
	if (mySel.contents == "") {
		// alert("선택된 텍스트가 없습니다.");
		panel_01.text = "ID : None";
	} else {
		mID = mySel.appliedConditions[0].name
		panel_01.text = "ID : " + mID;
	}
}

btn_10.onClick = function () {
	removeUnusedConditions();
}

btn_11.onClick = function () {
	var mySelID = app.selection[0].appliedConditions[0];

	if (mySelID == undefined) {
		alert("선택한 텍스트에 ID가 적용되어 있지 않습니다.");
		state = 0;
		panel_01.text = "None";
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
	panel_01.text = myFound.length + "(" + (state + 1) + ") : " + mySelID.name;

	// $.writeln(myFound.length);
	if (myFound.length == 1) {
		alert("선택한 OSD와 같은 ID를 사용하는 텍스트가 없습니다.");
		state = 0;
		exit();
	} else {
		myFound[state].select();
		app.activeWindow.activePage = myFound[state].parentTextFrames[0].parentPage;
	}
	state ++;
	if (state == myFound.length) state = 0;
}

btn_12.onClick = function () {
	eXportOSD();
}