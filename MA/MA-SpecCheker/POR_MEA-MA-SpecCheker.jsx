#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapUp.jsx";
#include "_reference/changeSemi-CapDown.jsx";

var w = new Window ('palette {text: "POR_MEA 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_03 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_02 = group02.add("button", [0, 0, 150, 30], "4G/LTE 병행 표기");
		btn_04 = group02.add("button", [0, 0, 150, 30], "표지 언어명 변경");

//----------------------------------------------------------------------------------------------------
btn_03.onClick = function () {
	RemoveBlankPages();
}
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
			app.findTextPreferences.findWhat = "Retirar a bateria";
			app.findTextPreferences.appliedParagraphStyle = "Heading1";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("배터리 분리 방법 설명글을 다음의 페이지에서 찾았습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name);
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {
				continue;
			}
		}
	alert("배터리 분리 방법 설명글을 찾을 수 없습니다.");
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	}
}
//----------------------------------------------------------------------------------------------------
btn_02.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findChangeTextOptions.caseSensitive = true;
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "4G/LTE";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("4G/LTE 병행 표기된 부분을 찾았습니다.");
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {
				continue
			}
		} catch(e) {}
	} alert ("4G/LTE 병행 표기된 부분을 찾을 수 없습니다. 문서를 확인한 후 수정하세요.");
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	var dest_file = "001_Cover.indd";
	for (var i=0; i < doc.length; i++) {
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Portuguese (MEA)";
				var founds00 = app.activeDocument.findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("표지 언어명이 이미 적용되어 있습니다.");
					exit ();
				} else {}
			} catch(e) {}

			//alert ("찾았다!");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Portuguese";
			app.findTextPreferences.appliedParagraphStyle = "Description";
			var founds = app.activeDocument.findText();
			if (founds.length < 1) {
				alert ("표지에 Portguese 표기된 위치를 찾을 수 없습니다.");
			} else
			founds[0].insertionPoints[-1].contents= " (MEA)";
			alert ("표지 언어명을 수정했습니다. 문서를 확인하세요.");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		} else {
			continue;
		}
	} 
}
//----------------------------------------------------------------------------------------------------
w.show();