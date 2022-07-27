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
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "la información IMEI";
			var foundsXX = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<foundsXX.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				gogo();
			} else {
				alert("IMEI 관련 문구가 없습니다.")
			}
		} catch(e) {}
	}
	function gogo() {
		for (var i=0; i < doc.length; i++) {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Nieprawidłowe podłączenie ładowarki może spowodować poważne uszkodzenie urządzenia. Żadne uszkodzenia wynikające z nieprawidłowej obsługi nie są objęte gwarancją.";
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
		var source_file = File("C:/MA-WikiSpec/POL/MA_online_UM_Pol_Battery_Polimer_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Pol_Battery_Polimer_spec.indd");
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