#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapDown.jsx";
#include "_reference/changeSemi-CapDown.jsx";

var w = new Window ('palette {text: "FRE_MEA 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "언어 표기명 변경");
		btn_04 = group02.add("button", [0, 0, 150, 30], "콜론 뒤 소문자 적용");
	group03 = w.add ('group {orientation: "column"}');
	group03.orientation = "row";
		btn_05 = group03.add("button", [0, 0, 150, 30], "세미콜론 뒤 소문자 적용");
		btn_06 = group03.add("button", [0, 0, 150, 30], "콜론 앞뒤 공백 확인");
	group05 = w.add ('group {orientation: "column"}');
	group05.orientation = "row";
		btn_07 = group05.add("button", [0, 0, 150, 30], "그린 이슈 + EN60950");
		btn_08 = group05.add("button", [0, 0, 150, 30], "라디오 앱 삭제");

//----------------------------------------------------------------------------------------------------
btn_01.onClick = function () {
	if(app.documents.length < 1){
		alert("인디자인 문서를 여세요!");
		exit();
		} else {
		var doc = app.documents;
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		for (var i=0; i < doc.length; i++){
			//app.activeDocument = doc[i];
			app.findTextPreferences.findWhat = "Retirer la batterie";
			app.findTextPreferences.appliedParagraphStyle = "Heading1";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("배터리 분리 방법 설명글을 다음의 페이지에서 찾았습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name);
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
				finds[0].select();
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				alert("배터리 분리 방법 설명글을 삭제해주세요.");
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
	var counter = 0
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeTextOptions.includeFootnotes = false;
		app.findChangeTextOptions.includeHiddenLayers = false;
		app.findChangeTextOptions.includeLockedLayersForFind = false;
		app.findChangeTextOptions.includeLockedStoriesForFind = false;
		app.findChangeTextOptions.includeMasterPages = false;

		app.findTextPreferences.findWhat = "French";
		app.changeTextPreferences.changeTo = "French (MEA)";
		var findText = doc[i].changeText();	
		for (var j=0; j<findText.length; j++) {
			findText[j];
			counter++
		} 
	} 
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
	alert (counter + " 개의 단어를 수정했습니다.");
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	changecaseDown();
}
//----------------------------------------------------------------------------------------------------
btn_05.onClick = function () {
	changeSemicaseDown();
}
//----------------------------------------------------------------------------------------------------
btn_06.onClick = function () {
	var doc = app.documents;
	wScontent = String.fromCharCode(0xA0);
	for (var i=0; i < doc.length; i++){
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeGrepOptions.includeFootnotes = false;
		app.findChangeGrepOptions.includeHiddenLayers = false;
		app.findChangeGrepOptions.includeLockedLayersForFind = false;
		app.findChangeGrepOptions.includeLockedStoriesForFind = false;
		app.findChangeGrepOptions.includeMasterPages = false;

		app.findGrepPreferences.findWhat = "(\\w+)(:)(\\w+)";
		app.changeGrepPreferences.changeTo = "$1" + wScontent + "$2 $3";
		doc[i].changeGrep();
	}
	for (var i=0; i < doc.length; i++){
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeGrepOptions.includeFootnotes = false;
		app.findChangeGrepOptions.includeHiddenLayers = false;
		app.findChangeGrepOptions.includeLockedLayersForFind = false;
		app.findChangeGrepOptions.includeLockedStoriesForFind = false;
		app.findChangeGrepOptions.includeMasterPages = false;

		app.findGrepPreferences.findWhat = "(\\w+)(:) (\\w+)";
		app.changeGrepPreferences.changeTo = "$1" + wScontent + "$2 $3";
		doc[i].changeGrep();
	}
	for (var i=0; i < doc.length; i++){
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeGrepOptions.includeFootnotes = false;
		app.findChangeGrepOptions.includeHiddenLayers = false;
		app.findChangeGrepOptions.includeLockedLayersForFind = false;
		app.findChangeGrepOptions.includeLockedStoriesForFind = false;
		app.findChangeGrepOptions.includeMasterPages = false;

		app.findGrepPreferences.findWhat = "(\\w+) (:)(\\w+)";
		app.changeGrepPreferences.changeTo = "$1" + wScontent + "$2 $3";
		doc[i].changeGrep();
	}
	for (var i=0; i < doc.length; i++){
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
		
		//Set the find options.
		app.findChangeGrepOptions.includeFootnotes = false;
		app.findChangeGrepOptions.includeHiddenLayers = false;
		app.findChangeGrepOptions.includeLockedLayersForFind = false;
		app.findChangeGrepOptions.includeLockedStoriesForFind = false;
		app.findChangeGrepOptions.includeMasterPages = false;

		app.findGrepPreferences.findWhat = "(\\w+)(:)(\\r)";
		app.changeGrepPreferences.changeTo = "$1" + wScontent + "$2$3";
		doc[i].changeGrep();

		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
	}
	alert ("변경 완료")
}
//----------------------------------------------------------------------------------------------------
btn_07.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		for (var i=0; i < doc.length; i++){
			//app.activeDocument = doc[i];
			app.findTextPreferences.findWhat = "Pour économiser l’énergie, débranchez le chargeur lorsque vous ne l’utilisez pas. Le chargeur n’étant pas muni d’une touche Marche/Arrêt, vous devez le débrancher de la prise de courant pour couper l’alimentation. Le socle de prise de courant doit être installé à proximité du matériel et doit être aisément accessible.";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("배터리 관련 문구를 찾았습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name + "\r문구를 교체합니다.");
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				finds[0].select();
				main();
			}
			else {
				continue;
			}
		}
		alert("배터리 관련 문구를 찾을 수 없습니다.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	}
	function main() {
		var destination_doc = app.activeDocument;
		//템플릿 파일 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var source_file = File("C:/MA-WikiSpec/FRE/MA_online_UM_Fre_Greenissue_Adaptor_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Fre_Greenissue_Adaptor_spec.indd");
		var sourcePages = source_doc.pages[0];
		var frame = sourcePages.textFrames[0];
		var selection = frame.texts.everyItem().select();
		app.copy();
		source_doc.close();
		destination_doc;
		app.paste();
		alert("그린 이슈 문구 및 Adaptor 관련 EN60950 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_08.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		var doc = app.documents;
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		for (var i=0; i < doc.length; i++){
			//app.activeDocument = doc[i];
			app.findTextPreferences.findWhat = "Écouter la radio FM";
			app.findTextPreferences.appliedParagraphStyle = "Heading2-APPLINK"
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("라디오 앱 내용이 있습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name + "\r관련 내용을 삭제해주세요.");
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				finds[0].select();
				exit();
			}
			else {
				continue;
			}
		}
		alert("라디오 관련 내용을 찾을 수 없습니다.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	}
}
//----------------------------------------------------------------------------------------------------
w.show();