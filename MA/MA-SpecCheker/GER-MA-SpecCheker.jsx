#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";

var w = new Window ('palette {text: "GER 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "과충전 주의 문구 추가");
		btn_04 = group02.add("button", [0, 0, 150, 30], "-");

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
			app.findTextPreferences.findWhat = "Akku entfernen";
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
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		for (var i=0; i < doc.length; i++){
			//app.activeDocument = doc[i];
			app.findTextPreferences.findWhat = "Laden Sie das Gerät nicht länger als eine Woche auf, da eine Überladung die Akkulebensdauer beeinträchtigen kann.";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("과충전 주의 문구가 있습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name);
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
		alert("과충전 주의 문구가 없습니다. 문구를 추가합니다.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	}
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		for (var i=0; i < doc.length; i++){
			app.findTextPreferences.findWhat = "Verwenden Sie ausschließlich das im Lieferumfang des Geräts enthaltene USB-Kabel vom Typ C. Bei der Verwendung von Micro-USB-Kabeln kann es zu Schäden am Gerät kommen."
			var finds = doc[i].findText();

			if (finds.length > 0) {
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				finds[0].select();
				finds[0].insertionPoints[-1].contents= "\r######";
				// alert("찾았다");
				main();
			}
			else {
				continue;
			}
		}
	} alert("과충전 주의 문구를 추가할 위치를 찾지 못했습니다.");

	function main() {
	var destination_doc = app.activeDocument;
	//템플릿 파일 열기
	// turn off warnings: missing fonts, links, etc.
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	var source_file = File("C:/MA-WikiSpec/GER/MA_online_UM_Ger_Charging_spec.indd"); // 사양문장 템플릿 문서 열기
	app.open(source_file);
	var source_doc = app.documents.item("MA_online_UM_Ger_Charging_spec.indd");
	var sourcePages = source_doc.pages[0];
	var frame = sourcePages.textFrames[0];
	var selection = frame.texts.everyItem().select();
	app.copy();
	source_doc.close();

	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	app.findTextPreferences.findWhat = "######";
	var foundSharp = app.activeDocument.findText();
	foundSharp[0].select();

	destination_doc;
	app.paste();
	alert("과충전 주의 문구를 추가했습니다. 문서를 확인하세요.");
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
}
//----------------------------------------------------------------------------------------------------
w.show();