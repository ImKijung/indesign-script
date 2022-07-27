#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapUp.jsx";

var w = new Window ('palette {text: "SPA_LTN 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_03 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 삭제");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_02 = group02.add("button", [0, 0, 150, 30], "표지 언어명 변경");
		btn_04 = group02.add("button", [0, 0, 150, 30], "Nano-SIM 카드 주의 확인");
	group03 = w.add ('group {orientation: "column"}');
	group03.orientation = "row";
		btn_05 = group03.add("button", [0, 0, 150, 30], "IMEI 정보 확인 방법 추가");
		btn_06 = group03.add("button", [0, 0, 150, 30], "Radio 기능 유무 확인");

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
			app.findTextPreferences.findWhat = "Retirar la batería";
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
//----------------------------------------------------------------------------------------------------
btn_02.onClick = function () {
	var doc = app.documents;
	var dest_file = "001_Cover.indd";
	for (var i=0; i < doc.length; i++) {
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Spanish (LTN)";
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
			app.findTextPreferences.findWhat = "Spanish";
			app.findTextPreferences.appliedParagraphStyle = "Description";
			var founds = app.activeDocument.findText();
			if (founds.length < 1) {
				alert ("표지에 Spanish 표기된 위치를 찾을 수 없습니다.");
			} else
			founds[0].insertionPoints[-1].contents= " (LTN)";
			alert ("표지 언어명을 수정했습니다. 문서를 확인하세요.");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		} else {
			continue;
		}
	} 
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {//Nano-SIM 카드 주의 문구 
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Utilice solo tarjetas nano-SIM.";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("Nano-SIM 카드 주의 문구가 이미 있습니다.");
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {}
		} catch(e) {}
	}
	alert ("Nano-SIM 카드 주의 문구를 찾지 못했습니다.");
}
//----------------------------------------------------------------------------------------------------
btn_05.onClick = function () {//IMEI 정보 확인 방법 추가
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Comprobar la información IMEI del dispositivo";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("IMEI 정보 확인 방법 문구가 이미 적용되어 있습니다.");
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {}
		} catch(e) {}
	}
	var counter00 = 0;
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "la información IMEI";
		var foundsXX = doc[i].findText();
		for (var j=0; j<foundsXX.length; j++) {
			counter00 ++;
		}
	}
	if (counter00 > 0) {
		gogo();
	} else {
		alert("IMEI 관련 문구가 없습니다.")
	}
	function gogo() {
		for (var i=0; i < doc.length; i++) {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Información de la batería: Permite visualizar la información y el estado de la batería del dispositivo.";
			var founds = doc[i].findText();
			if (founds.length > 0) {
				app.activeDocument = doc[i];
				founds[0].insertionPoints[-1].contents= "\r######";
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				main();
			} else {}
		}
		alert ("사양을 넣을 위치를 찾지 못했습니다.");
	}

	function main() {
		var destination_doc = app.activeDocument;
		//템플릿 파일 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var source_file = File("C:/MA-WikiSpec/SPA_LTN/MA_online_UM_Spa_LTN_IMEI_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Spa_LTN_IMEI_spec.indd");
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
		alert("IMEI 정보 확인 방법 문구를 적용했습니다. 문서를 확인해주세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_06.onClick = function () {
	var doc = app.documents;
	var counter = 0;
	for (var i=0; i < doc.length; i++) {
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "Radio";
		app.findTextPreferences.appliedParagraphStyle = "Heading1-APPLINK-Nosub"
		var foundsXX = doc[i].findText();
		for (var j=0; j<foundsXX.length; j++) {
			counter ++;
		}
	}
	//alert(counter);
	if (counter > 0) {
		alert("Radio 앱이 적용된 매뉴얼입니다.")
	} else {
		alert("Radio 앱이 적용되지 않았습니다.")
	}
}
//----------------------------------------------------------------------------------------------------
w.show();