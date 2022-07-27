#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";

var w = new Window ('palette {text: "ARA 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "영단어 LtR 스타일 적용");
		btn_04 = group02.add("button", [0, 0, 150, 30], "화살표 방향 바꾸기");

//----------------------------------------------------------------------------------------------------
btn_01.onClick = function specChecker01() {
	if(app.documents.length < 1){
		alert("인디자인 문서를 여세요!");
		exit();
		} else {
		var doc = app.documents;
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		for (var i=0; i < doc.length; i++){
			//app.activeDocument = doc[i];
			app.findTextPreferences.findWhat = "إزالة البطارية";
			app.findTextPreferences.appliedParagraphStyle = "Heading1";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("배터리 분리 방법 설명글을 다음의 페이지에서 찾았습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name + "\r배터리 분리 방법 설명글을 삭제하세요.");
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
				finds[0].select();
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			}
			else {
				continue;
			}
		}
		alert("배터리 분리 방법 설명글을 찾을 수 없습니다.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		}
}
//----------------------------------------------------------------------------------------------------
btn_02.onClick = function () {
	RemoveBlankPages();
}
//----------------------------------------------------------------------------------------------------
btn_03.onClick = function () {
	var docs = app.documents;
	var dest_file = "007_Copyright.indd";
	for (var i=0; i < docs.length; i++){
		if (docs[i].name == dest_file) {
			app.activeDocument = docs[i];
		}
	}
	var rtlcode = String.fromCharCode(0x200F);
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	//Set the find options.
	app.findChangeTextOptions.includeFootnotes = false;
	app.findChangeTextOptions.includeHiddenLayers = false;
	app.findChangeTextOptions.includeLockedLayersForFind = false;
	app.findChangeTextOptions.includeLockedStoriesForFind = false;
	app.findChangeTextOptions.includeMasterPages = false;

	app.findTextPreferences.findWhat = "Wi-Fi®";
	var found01 = app.activeDocument.findText();
	if (found01.length > 0) {
		found01[0].insertionPoints[-1].contents= rtlcode;
	} else {}

	app.findTextPreferences.findWhat = "Wi-Fi Direct™";
	var found02 = app.activeDocument.findText();
	if (found02.length > 0) {
		found02[0].insertionPoints[-1].contents= rtlcode;
	} else {}

	app.findTextPreferences.findWhat = "Wi-Fi CERTIFIED™";
	var found03 = app.activeDocument.findText();
	if (found03.length > 0) {
		found03[0].insertionPoints[-1].contents= rtlcode;
	} else {}

	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	for (var i=0;i<docs.length;i++){
		var openDocs = app.documents.everyItem().getElements();
		app.activeDocument = openDocs[openDocs.length-1];
		applyLtR();
	} 
	alert("Left to Right 적용 완료");

	function applyLtR () {
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;

		//Set the find options.
		app.findChangeGrepOptions.includeFootnotes = false;
		app.findChangeGrepOptions.includeHiddenLayers = false;
		app.findChangeGrepOptions.includeLockedLayersForFind = false;
		app.findChangeGrepOptions.includeLockedStoriesForFind = false;
		app.findChangeGrepOptions.includeMasterPages = false;

		var greps = [
			{"findWhat":"Samsung account"},
			{"findWhat":"MirrorLink"},
			{"findWhat":"Dolby Atmos"},
			{"findWhat":"Smart Lock"},
			{"findWhat":"Always On Display"},
			{"findWhat":"FaceWidgets"},
			{"findWhat":"Google Play"},
			{"findWhat":"Samsung Pass"},
			{"findWhat":"Secure Wi-Fi"},
			{"findWhat":"Galaxy Store"},
			{"findWhat":"Samsung Cloud"},
			{"findWhat":"Secure Folder"},
			{"findWhat":"Samsung Notes"},
			{"findWhat":"Smart Switch"},
			{"findWhat":"Google account"},
			{"findWhat":"Bixby Routines"},
			{"findWhat":"Game Launcher"},
			{"findWhat":"Dual Messenger"},
			{"findWhat":"Voice Assistant"},
			{"findWhat":"Bixby Vision"},
			{"findWhat":"Bixby Home"},
			{"findWhat":"Samsung Pay"},
			{"findWhat":"Samsung Health"},
			{"findWhat":"Galaxy Wearable"},
			{"findWhat":"Samsung Members"},
			{"findWhat":"screen mirroring "},
			{"findWhat":"Bluetooth®"},
			{"findWhat":"Wi-Fi®"},
			{"findWhat":"Wi-Fi Direct™"},
			{"findWhat":"Wi-Fi CERTIFIED™"},
		]
		for(var m=0; m<greps.length; m++) {
			app.findGrepPreferences.findWhat = greps[m].findWhat;
			//app.changeGrepPreferences.changeTo = greps[m].changeTo;
			var found = app.activeDocument.findGrep();
			for (var i=0;i<found.length;i++) {
				found[i].select();
				var selection = app.selection[0];
				selection.texts.everyItem().characterDirection = CharacterDirectionOptions.LEFT_TO_RIGHT_DIRECTION;
				selection.texts.everyItem().noBreak = true;
			}
		}
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
	}
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	var arcounter = 0
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeTextOptions.includeFootnotes = false;
		app.findChangeTextOptions.includeHiddenLayers = false;
		app.findChangeTextOptions.includeLockedLayersForFind = false;
		app.findChangeTextOptions.includeLockedStoriesForFind = false;
		app.findChangeTextOptions.includeMasterPages = false;

		app.findTextPreferences.findWhat = "→";
		app.findTextPreferences.appliedCharacterStyle = "C_SingleStep";
		app.changeTextPreferences.changeTo = "←";
		var findText = doc[i].changeText();	
		for (var j=0; j<findText.length; j++) {
			findText[j];
			arcounter++
		} 
	}
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
	alert (arcounter + " 개의 화살표를 수정했습니다.");

	var trcounter = 0
	for (var j=0; j < doc.length; j++){
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeTextOptions.includeFootnotes = false;
		app.findChangeTextOptions.includeHiddenLayers = false;
		app.findChangeTextOptions.includeLockedLayersForFind = false;
		app.findChangeTextOptions.includeLockedStoriesForFind = false;
		app.findChangeTextOptions.includeMasterPages = false;

		app.findTextPreferences.findWhat = "►";
		app.findTextPreferences.appliedCharacterStyle = "C_SingleStep";
		app.changeTextPreferences.changeTo = "◄";
		var findText = doc[j].changeText();	
		for (var k=0; k<findText.length; k++) {
			findText[k];
			trcounter++
		} 
	}
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
	alert (trcounter + " 개의 세모를 수정했습니다.");
}
//----------------------------------------------------------------------------------------------------
w.show();