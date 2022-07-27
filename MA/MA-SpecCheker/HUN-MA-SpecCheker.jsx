#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapUp.jsx";

var w = new Window ('palette {text: "HUN 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "USIM 단어/문장 삭제");
		btn_04 = group02.add("button", [0, 0, 150, 30], "별매품 단어 삭제 및 변경");

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
			app.findTextPreferences.findWhat = "Az akkumulátor eltávolítása";
			app.findTextPreferences.appliedParagraphStyle = "Heading1";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("배터리 분리 방법 설명글을 다음의 페이지에서 찾았습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name);
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
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
	var doc = app.documents;
	var counter = 0;
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeTextOptions.includeFootnotes = false;
		app.findChangeTextOptions.includeHiddenLayers = false;
		app.findChangeTextOptions.includeLockedLayersForFind = false;
		app.findChangeTextOptions.includeLockedStoriesForFind = false;
		app.findChangeTextOptions.includeMasterPages = false;

		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;

		app.findTextPreferences.findWhat = "UMTS- és HSDPA- (3G és 3G+) szolgáltatásokhoz USIM- (Universal Subscriber Identity Module – általános előfizetőazonosító modul) kártya is beszerezhető.\r";
		app.changeTextPreferences.changeTo = "";
		var found03 = doc[i].changeText();
		for (var n=0;n<found03.length;n++) {
			found03[n];
			counter++;
		}

		app.findTextPreferences.findWhat = "-vagy USIM-";
		app.changeTextPreferences.changeTo = "";
		var found01 = doc[i].changeText();
		for (var j=0;j<found01.length;j++) {
			found01[j];
			counter++;
		}
		
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;

		app.findTextPreferences.findWhat = "USIM";
		app.changeTextPreferences.changeTo = "";
		var found02 = doc[i].changeText();
		for (var k=0;k<found02.length;k++) {
			found02[k];
			counter++;
		}
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
	}
	alert ("USIM 관련 단어 및 문구를 " + counter + " 개 발견해 삭제했습니다.");
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	// var text01 = "(KÜLÖN KAPHATÓK)";
	var text02 = "egyes (külön kapható)";
	// var chgText = "mágneseket tartalmazó";
	// var counter01 = 0;
	var counter02 = 0;
	for (var i=0; i < doc.length; i++){
		//Set the find options.
		app.findChangeTextOptions.includeFootnotes = false;
		app.findChangeTextOptions.includeHiddenLayers = false;
		app.findChangeTextOptions.includeLockedLayersForFind = false;
		app.findChangeTextOptions.includeLockedStoriesForFind = false;
		app.findChangeTextOptions.includeMasterPages = false;

		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;

		app.findTextPreferences.findWhat = text02;
		// app.changeTextPreferences.changeTo = chgText;
		var found02 = doc[i].findText();
		for (var j=0;j<found02.length;j++) {
			found02[j];
			counter02++;
			if (found02.length > 0) { // something has been found
				alert("자석문구 내 별매품(sold separately) 단어가 있습니다.\r해당 위치로 이동합니다. 페이지: " + found02[0].parentTextFrames[0].parentPage.name);
				app.activeDocument = doc[i];
				app.activeWindow.activePage = found02[0].parentTextFrames[0].parentPage;
				found02[0].select();
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			}
			else {
				continue;
			}
		} 
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
	}
	alert(text02 + " 문구가 문서 전체에 없습니다.");
}
//----------------------------------------------------------------------------------------------------
w.show();