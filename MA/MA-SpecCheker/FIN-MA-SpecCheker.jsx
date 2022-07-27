#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapUp.jsx";

var w = new Window ('palette {text: "FIN 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "콜론 뒤 대문자 적용");
		btn_04 = group02.add("button", [0, 0, 150, 30], "법인 주소 확인");

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
			app.findTextPreferences.findWhat = "Akun irrottaminen";
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
	changecaseUp();
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	var dest_file = "007_Copyright.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Jos tarvitset apua tai on kysyttävää, käy sivullamme";
				var founds00 = app.activeDocument.findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("법인 주소가 이미 적용되어 있습니다.");
					exit ();
				} else {}
			} catch(e) {}

			//alert ("찾았다!");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Tekijänoikeus";
			app.findTextPreferences.appliedParagraphStyle = "Heading1-H3";
			var founds = app.activeDocument.findText();
			founds[0].insertionPoints[0].contents= "######\r";
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			main();
		} else {
			continue;
		}
	} alert (dest_file + " 문서가 없습니다.");

	function main() {
		var destination_doc = app.activeDocument;
		//템플릿 파일 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var source_file = File("C:/MA-WikiSpec/FIN/MA_online_UM_Fin_CorporateAddress_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Fin_CorporateAddress_spec.indd");
		var sourcePages = source_doc.pages[0];
		var frame = sourcePages.textFrames[0];
		var selection = frame.texts.everyItem().select();
		app.copy();
		source_doc.close();
		
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "######";
		app.findTextPreferences.appliedParagraphStyle = "Heading1-H3";
		var foundSharp = app.activeDocument.findText();
		foundSharp[0].select();
		
		destination_doc;
		app.paste();
		alert("법인 주소를 추가했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
w.show();