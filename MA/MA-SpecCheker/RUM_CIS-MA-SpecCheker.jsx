#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapUp.jsx";

var w = new Window ('palette {text: "RUM_CIS 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_03 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 삭제");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_02 = group02.add("button", [0, 0, 150, 30], "Dual SIM 과금 문구 추가");
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
			app.findTextPreferences.findWhat = "Scoaterea bateriei";
			app.findTextPreferences.appliedParagraphStyle = "Heading1";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("배터리 분리 방법 설명글을 다음의 페이지에서 찾았습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name + "\r배터리 분리 방법 설명글을 삭제하세요.");
				finds[0].select();
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
btn_02.onClick = function () {//Dual SIM 과금 문구 추가
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "În funcție de regiune sau de furnizorii de servicii, redirecționarea apelurilor în convorbire de pe un SIM pe altul este posibil să nu funcționeze. Verificați disponibilitatea serviciului cu operatorii de telefonie mobilă. Utilizarea acestei funcții poate genera costuri adiționale.";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("Dual SIM 과금 문구가 이미 있습니다.");
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {}
		} catch(e) {}
	}
	for (var i=0; i < doc.length; i++) {
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "Dacă introduceți două cartele SIM sau USIM, puteți avea două numere de telefon sau doi furnizori de servicii la un singur dispozitiv.";
		var founds = doc[i].findText();
		if (founds.length > 0) {
			app.activeDocument = doc[i];
			founds[0].insertionPoints[-1].contents= "\r######";
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			main();
		} else {}
	}
	alert ("사양을 넣을 위치를 찾지 못했습니다.");

	function main() {
		var destination_doc = app.activeDocument;
		//템플릿 파일 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var source_file = File("C:/MA-WikiSpec/RUM_CIS/MA_online_UM_Rum_CIS_DualSIM_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Rum_CIS_DualSIM_spec.indd");
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
		alert("Dual SIM 과금 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
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
				app.findTextPreferences.findWhat = "Romanian (CIS)";
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
			app.findTextPreferences.findWhat = "Romanian";
			app.findTextPreferences.appliedParagraphStyle = "Description";
			var founds = app.activeDocument.findText();
			if (founds.length < 1) {
				alert ("표지에 Romanian 표기된 위치를 찾을 수 없습니다.");
			} else
			founds[0].insertionPoints[-1].contents= " (CIS)";
			alert ("표지 언어명을 수정했습니다. 문서를 확인하세요.");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		} else {
			continue;
		}
	} 
}
//----------------------------------------------------------------------------------------------------
w.show();