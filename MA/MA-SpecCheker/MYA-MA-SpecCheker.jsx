#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";

var w = new Window ('palette {text: "MYA 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 삭제");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "콜론(:) 자국어 표기");
		btn_04 = group02.add("button", [0, 0, 150, 30], "마침표(.) 자국어 표기");

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
			app.findTextPreferences.findWhat = "ဘက္ထရီ ဖယ္ရွားျခင္း";
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
	var counter = 0;
	for (var j=0; j< doc.length; j++) {
		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;

		app.findGrepPreferences.findWhat = "\\:";
		app.changeGrepPreferences.changeTo = "\\x{104B}\\s\\x{104B}";

		var myFound = doc[j].changeGrep();

		for (var i=0; i < myFound.length; i++) {
			myFound[i];
			counter ++;
		} 
	} 
	alert (counter + " 개의 콜론(:)을 수정했습니다.");
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	var counter = 0;
	for (var j=0; j< doc.length; j++) {
		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;

		app.findGrepPreferences.findWhat = "([\\x{1000}-\\x{109F}])\\.";
		app.changeGrepPreferences.changeTo = "$1\\x{104B}";

		var myFound = doc[j].changeGrep();

		for (var i=0; i < myFound.length; i++) {
			myFound[i];
			counter ++;
		} 
	} 
	alert (counter + " 개의 마침표(.)를 수정했습니다.");
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
}
//----------------------------------------------------------------------------------------------------
w.show();