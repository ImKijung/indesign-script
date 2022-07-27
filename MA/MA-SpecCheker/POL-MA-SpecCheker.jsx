#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapDown.jsx";

var w = new Window ('palette {text: "POL 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "콜론(:) 뒤 소문자 변경");
		btn_04 = group02.add("button", [0, 0, 150, 30], "samsung.pl 변경");
	group03 = w.add ('group {orientation: "column"}');
	group03.orientation = "row";
		btn_05 = group03.add("button", [0, 0, 150, 30], "인터넷 과금 문구 적용");
		btn_06 = group03.add("button", [0, 0, 150, 30], "Headset Jack 수치 삭제");
	group04 = w.add ('group {orientation: "column"}');
	group04.orientation = "row";
		btn_07 = group04.add("button", [0, 0, 150, 30], "MMS 관련 노트 문구 적용");
		btn_08 = group04.add("button", [0, 0, 150, 30], "배터리 용량 적용(신규)");
	group05 = w.add ('group {orientation: "column"}');
	group05.orientation = "row";
		btn_09 = group05.add("button", [0, 0, 150, 30], "리튬이온 배터리 적용");
		btn_10 = group05.add("button", [0, 0, 150, 30], "리튬폴리머 배터리 적용");
	group06 = w.add ('group {orientation: "column"}');
	group06.orientation = "row";
		btn_11 = group06.add("button", [0, 0, 150, 30], "국가별 상이 문구 적용");
		btn_12 = group06.add("button", [0, 0, 150, 30], "4G 아이콘 삭제");
	group07 = w.add ('group {orientation: "column"}');
	group07.orientation = "row";
		btn_13 = group07.add("button", [0, 0, 150, 30], "Speed dial 부분 숫자 상이");
		btn_14 = group07.add("button", [0, 0, 150, 30], "-");

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
			app.findTextPreferences.findWhat = "Wyjmowanie baterii";
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
	changecaseDown ();
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	var count = 0;
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "www.samsung.com";
		app.findTextPreferences.appliedCharacterStyle = "C_URL";
		app.changeTextPreferences.changeTo = "www.samsung.pl";
		var founds = doc[i].changeText();
		for (var n=0; n<founds.length; n++) {
			founds[n];
			count++;
		}
		app.findTextPreferences.appliedCharacterStyle = "C_URL-Important";
		var founds2 = doc[i].changeText();
		for (var n=0; n<founds2.length; n++) {
			founds2[n];
			count++;
		}
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "www.samsung.pl/smartswitch";
		app.findTextPreferences.appliedCharacterStyle = "C_URL";
		app.changeTextPreferences.changeTo = "www.samsung.com/smartswitch";
		doc[i].changeText();
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "www.samsung.pl/samsungpay";
		app.findTextPreferences.appliedCharacterStyle = "C_URL";
		app.changeTextPreferences.changeTo = "www.samsung.com/samsungpay";
		doc[i].changeText();
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "www.samsung.pl/samsung-health";
		app.findTextPreferences.appliedCharacterStyle = "C_URL";
		app.changeTextPreferences.changeTo = "www.samsung.com/samsung-health";
		doc[i].changeText();
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "www.samsung.pl/global/ecodesign_energy";
		app.findTextPreferences.appliedCharacterStyle = "C_URL";
		app.changeTextPreferences.changeTo = "www.samsung.com/global/ecodesign_energy";
		doc[i].changeText();
	}
	alert("웹주소를 찾아 변경했습니다.");
	
	var dest_file = "001_Cover.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "www.samsung.pl^pwww.samsung.com";
			app.findTextPreferences.appliedParagraphStyle = "Description-Cover";
			var foundPl = app.activeDocument.findText();
			var countPl = 0;
			for (var j=0; j<foundPl.length; j++) {
				foundPl[j];
				countPl++;
			}
			if (countPl == 0) {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "www.samsung.pl";
				app.findTextPreferences.appliedParagraphStyle = "Description-Cover";
				app.changeTextPreferences.changeTo = "www.samsung.pl^pwww.samsung.com"
				app.activeDocument.changeText();
				alert ("커버의 웹주소를 변경했습니다.");
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {
				alert ("커버에 변경할 부분을 찾지 못했습니다. 웹주소 변경을 중지합니다.")
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			}
		} else {
			continue
		}
	} alert (dest_file + " 파일이 열려있지 않습니다. 파일을 열고 다시 스크립트를 실행하세요.")
}
//----------------------------------------------------------------------------------------------------
btn_05.onClick = function () {
	var doc = app.documents;
	var countLTE = 0;
	for (var i=0; i < doc.length; i++){
		try {
			app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
			app.findGrepPreferences.findWhat = "UMTS|HSDPA|HSPA+|LTE";
			var foundLTE = doc[i].findGrep();
			for (var j=0; j < foundLTE.length; j++) {
				countLTE ++;
				if (countLTE > 0) {
					alert ("Wi-Fi 전용 모델이 아닙니다. 인터넷 과금 관련 문구를 추가합니다.");
					gogo();
				} else if (countLTE < 1) {
					alert ("Wi-Fi 전용 모델입니다.");
					exit();
				} else {}
			}
		} catch(e) {}
	}
	function gogo() {
		for (var i=0; i < doc.length; i++){
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Zalecamy korzystanie z dedykowanych taryf lub pakietów danych, które umożliwiają korzystanie z transmisji danych i pozwolą uniknąć dodatkowych kosztów z tym związanych. Włączony telefon, może być na stałe podłączony do Internetu i automatycznie synchronizować się z usługami opartymi na transmisji danych.";
				var founds00 = doc[i].findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("인터넷 과금 관련 문구가 이미 적용되어 있습니다.");
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
			app.findTextPreferences.findWhat = "Dotknij pola adresu.";
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
		var source_file = File("C:/MA-WikiSpec/POL/MA_online_UM_Pol_InternetBilling_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Pol_InternetBilling_spec.indd");
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
		alert("인터넷 과금 관련 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_06.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "Wygląd urządzenia";
		app.findTextPreferences.appliedParagraphStyle = "Heading2";
		var founds = doc[i].findText();
		if (founds.length > 0) { // something has been found
			app.activeDocument = doc[i];
			app.activeWindow.activePage = founds[0].parentTextFrames[0].parentPage;
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			founds[0].select();
			alert("Headset Jack 번역값이 수정되어 있는지 확인하세요.")
			exit();
		}
		else {
			continue;
		}
	}
}
//----------------------------------------------------------------------------------------------------
btn_07.onClick = function () {
	var doc = app.documents;
	var countLTE = 0;
	for (var i=0; i < doc.length; i++){
		try {
			app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
			app.findGrepPreferences.findWhat = "UMTS|HSDPA|HSPA+|LTE";
			var foundLTE = doc[i].findGrep();
			for (var j=0; j < foundLTE.length; j++) {
				countLTE ++;
				if (countLTE > 0) {
					alert ("Wi-Fi 전용 모델이 아닙니다. MMS 관련 노트를 추가합니다.");
					gogo();
				} else if (countLTE < 1) {
					alert ("Wi-Fi 전용 모델입니다.");
					exit();
				} else {}
			}
		} catch(e) {}
	}
	function gogo() {
		for (var i=0; i < doc.length; i++){
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Maksymalna dopuszczalna liczba znaków w wiadomości SMS zależy od operatora sieci. Jeżeli wiadomość przekroczy maksymalną liczbę znaków, urządzenie ją podzieli.^pMożesz wybrać rodzaj alfabetu dla nowych wiadomości SMS w Ustawieniach w menu Obsługiwane znaki. Po wybraniu opcji Automatyczny telefon zmieni kodowanie z alfabetu GSM na Unicode, jeśli zostanie wprowadzony znak Unicode. Użycie kodowania Unicode spowoduje zmniejszenie maksymalnej liczby znaków w wiadomości o około połowę.";
				var founds00 = doc[i].findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("MMS 관련 노트 문구가 이미 적용되어 있습니다.");
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
			app.findTextPreferences.findWhat = "Za wysyłanie wiadomości podczas korzystania z roamingu mogą być naliczane dodatkowe opłaty.";
			var founds = doc[i].findText();
			if (founds.length > 0) {
				app.activeDocument = doc[i];
				founds[0].select();
				app.selection[0].paragraphs[0].appliedParagraphStyle = app.activeDocument.paragraphStyles.item("UnorderList_1-Cell");
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
		var source_file = File("C:/MA-WikiSpec/POL/MA_online_UM_Pol_MMSNote_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Pol_MMSNote_spec.indd");
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
		alert("MMS 관련 노트 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}

btn_08.onClick = function () {//배터리 용량 문구 적용(신규)
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Typowa pojemność została sprawdzona w warunkach laboratoryjnych zapewnianych przez firmę zewnętrzną.";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("(신규)배터리 용량 관련 문구가 이미 적용되어 있습니다.");
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
		app.findTextPreferences.findWhat = "Używaj wyłącznie kabla z wtyczką USB typu C dostarczonego z urządzeniem. Użycie kabla micro USB może spowodować uszkodzenie urządzenia.";
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
		var source_file = File("C:/MA-WikiSpec/POL/MA_online_UM_Pol_Battery_New_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Pol_Battery_New_spec.indd");
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
		alert("(신규)배터리 용량 관련 문구를 적용했습니다. 배터리 용량을 확인하고 수정하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}

btn_09.onClick = function () {//리튬이온배터리
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Urządzenie posiada baterię litowo-jonową o pojemności";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("리튬-이온 배터리 용량 관련 문구가 이미 적용되어 있습니다.");
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

	function main() {
		var destination_doc = app.activeDocument;
		//템플릿 파일 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var source_file = File("C:/MA-WikiSpec/POL/MA_online_UM_Pol_Battery_Ion_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Pol_Battery_Ion_spec.indd");
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
		alert("리튬-이온 배터리 용량 관련 문구를 적용했습니다. 배터리 용량을 확인하고 수정하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}

btn_10.onClick = function () {//리튬폴리머배터리
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Urządzenie posiada baterię litowo-polimerową o pojemności";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("리튬-폴리머 배터리 용량 관련 문구가 이미 적용되어 있습니다.");
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
		alert("리튬-폴리머 배터리 용량 관련 문구를 적용했습니다. 배터리 용량을 확인하고 수정하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
btn_11.onClick = function () {//국가별 상이 문구 적용
	var doc = app.documents;
	var dest_file = "007_Copyright.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "W zależności od kraju i operatora, karty SIM, urządzenie i akcesoria mogą wyglądać inaczej niż na ilustracjach zamieszczonych w niniejszej instrukcji.";
				var founds00 = app.activeDocument.findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("국가별 상이 문구가 이미 적용되어 있습니다.");
					exit ();
				} else {}
			} catch(e) {}

			//alert ("찾았다!");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Prawa autorskie";
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
		var source_file = File("C:/MA-WikiSpec/POL/MA_online_UM_Pol_ByCountry_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Pol_ByCountry_spec.indd");
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
		alert("국가별 상이 문구를 추가했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}

btn_12.onClick = function () {//4G 아이콘 삭제
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Połączono z siecią LTE";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("폴란드는 4G 통신을 지원하지 않습니다. 4G 아이콘을 삭제하세요.");
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {}
		} catch(e) {}
	}
}

btn_13.onClick = function () {//Speed dial 부분 숫자 상이
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
			app.findGrepPreferences.findWhat = "W przypadku numerów szybkiego wybierania większych od \\d+ dotknij pierwszych cyfr numeru, a następnie dotknij ostatniej cyfry i przytrzymaj ją.";
			var founds00 = doc[i].findGrep();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				alert ("Speedi Dial 관련, 영문에는 10 으로 되어있으나 폴란드어 번역은 9가 맞습니다. 수치가 9인지 확인하세요.");
				app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
				exit();
			} else {}
		} catch(e) {}
	}
}
w.show();