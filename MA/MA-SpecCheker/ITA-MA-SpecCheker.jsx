#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapDown.jsx";

var w = new Window ('palette {text: "ITA 사양 Check"}');
	//w.main = w.add ('group');
	group01 = w.add ('group {orientation: "column"}');
	group01.orientation = "row";
		btn_02 = group01.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group01.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "콜론 뒤 소문자 적용");
		btn_04 = group02.add("button", [0, 0, 150, 30], "서비스/데이터 문구 추가");
	group03 = w.add ('group {orientation: "column"}');
	group03.orientation = "row";
		btn_05 = group03.add("button", [0, 0, 150, 30], "구성품 상이 문구 추가");
		btn_06 = group03.add("button", [0, 0, 150, 30], "Service Provider 삭제");
	group04 = w.add ('group {orientation: "column"}');
	group04.orientation = "row";
		btn_07 = group04.add("button", [0, 0, 150, 30], "상태 아이콘 문구 확인");
		btn_08 = group04.add("button", [0, 0, 150, 30], "Settings 챕터 번역 수정");
	group05 = w.add ('group {orientation: "column"}');
	group05.orientation = "row";
		btn_09 = group05.add("button", [0, 0, 150, 30], "삼성페이 문구 수정");
		btn_10 = group05.add("button", [0, 0, 150, 30], "-");

		
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
			app.findTextPreferences.findWhat = "Rimozione della batteria";
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
	changecaseDown();
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	var beforeText = "Questo dispositivo supporta servizi e applicazioni che potrebbero richiedere una connessione dati attiva per il loro funzionamento ed aggiornamento. Come impostazione predefinita, la connessione dati è sempre attiva su questo dispositivo. Verificate i costi di connessione con il vostro gestore telefonico. A seconda del gestore telefonico e del piano tariffario, alcuni servizi potrebbero non essere disponibili.^pPer disabilitare la connessione dati, nel menu Applicazioni, selezionate Impostaz. → Connessioni → Utilizzo dati e deselezionate Connessione dati.";

	var changeText = "Potete visualizzare le informazioni legali relative al dispositivo, in base al vostro paese. Per visualizzare le informazioni, avviate l'applicazione Impostaz. e toccate (.+)\\."
	// var changeText = "Potete visualizzare le informazioni legali relative al dispositivo, in base al vostro paese. Per visualizzare le informazioni, avviate l'applicazione Impostaz. e toccate Informazioni sul telefono → Regulatory information."
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = beforeText;
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("서비스/데이터 문구가 이미 추가되어 있습니다.");
				app.activeDocument = doc[i];
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				exit ();
			} else {
				app.findGrepPreferences.findWhat = changeText;
				var finds = doc[i].findGrep();
				if (finds.length > 0) { // something has been found
					//alert("찾았다");
					app.activeDocument = doc[i];
					app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
					//finds[0].select();
					finds[0].insertionPoints[-1].contents= "\r######";
					app.findGrepPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
					main();
				} else {
					continue;
				}
			}
		} catch(e) {}
	}
	// alert("교체할 문장을 찾지 못했습니다. 수동으로 확인하시기 바랍니다.");
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	alert("서비스/데이터 문구를 추가할 위치를 찾을 수 없습니다.\n스크립트 수정이 필요하니 개발자에게 문의하시기 바랍니다.")
	exit();

	function main() {
		var doc = app.activeDocument;
		//템플릿 파일 열기
		var destination_doc = app.activeDocument;
		var source_file = File("C:/MA-WikiSpec/ITA/MA_online_UM_Ita_service_data_spec.indd"); // 사양문장 템플릿 문서 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Ita_service_data_spec.indd");
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
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		alert("서비스/데이터 연결 및 연결 해제 문구를 교체했습니다. 문서를 확인하세요.");
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_05.onClick = function () {
	var doc = app.documents;
	var dest_file = "007_Copyright.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Il dispositivo e gli accessori illustrati in questo manuale potrebbero variare in base al Paese nel quale i prodotti vengono distribuiti.";
				var founds00 = app.activeDocument.findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("구성품 상이 문구가 이미 적용되어 있습니다.");
					exit ();
				} else {}
			} catch(e) {}
			//alert ("찾았다!");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Copyright";
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
		var source_file = File("C:/MA-WikiSpec/ITA/MA_online_UM_Ita_accessory_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Ita_accessory_spec.indd");
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
		alert("구성품 상이 문구를 추가했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_06.onClick = function () {
	var doc = app.documents;
	var dest_file = "003_Basics.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			//alert ("찾았다!");
			app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
			app.findGrepPreferences.findWhat = "EDGE|UMTS|HSDPA|HSPA+";
			var founds = doc[i].findGrep();
			var counter = 0
			for (var j=0; j < founds.length; j++) {
				founds[j].select();
				counter ++;
				} //alert (counter);
				if (counter > 0) {
					app.activeDocument = doc[i];
					app.activeWindow.activePage = founds[0].parentTextFrames[0].parentPage;
					alert ("Wi-Fi 전용 모델이 아닙니다.");
					app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
					exit();
				} else if (counter == 0) {
					alert ("Wi-Fi 전용 모델입니다. Service provider 문구를 삭제합니다.");
					app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
					delServiceProvider();
				}
		} else {
			continue;
		}
	} alert (dest_file + " 문서가 없습니다.");

	function delServiceProvider () {
		var doc = app.documents;
		for (var i=0; i < doc.length; i++){
			//alert (doc[i].name);
			app.findTextPreferences = NothingEnum.nothing;
			app.changeTextPreferences = NothingEnum.nothing;

			//Set the find options.
			app.findChangeTextOptions.includeFootnotes = false;
			app.findChangeTextOptions.includeHiddenLayers = false;
			app.findChangeTextOptions.includeLockedLayersForFind = false;
			app.findChangeTextOptions.includeLockedStoriesForFind = false;
			app.findChangeTextOptions.includeMasterPages = false;
			
			//^p 엔터가 있을 경우 추가
			var teXts = [
				{"findWhat":" al gestore telefonico,","changeTo":""},
				{"findWhat":" Per le applicazioni installate da voi, contattate il vostro gestore telefonico.","changeTo":""},
				{"findWhat":" o al gestore telefonico","changeTo":""},
				{"findWhat":" Per effettuare la registrazione o per ottenere maggiori informazioni sul servizio, contattate il vostro gestore telefonico.","changeTo":""},
				{"findWhat":"PUK: se la scheda SIM o USIM è bloccata, solitamente in seguito al ripetuto inserimento di un PIN errato. Dovete inserire il PUK fornito dal gestore telefonico.^p","changeTo":""},
				{"findWhat":"PIN2: quando accedete a un menu che lo richiede il PIN2, dovete inserire il PIN2 fornito con la scheda SIM o USIM. Per maggiori informazioni, rivolgetevi al vostro gestore telefonico.^p","changeTo":""},
				{"findWhat":"L'accesso ad alcune opzioni è soggetto a registrazione. Per maggiori informazioni, rivolgetevi al vostro gestore telefonico.^p","changeTo":""},
				{"findWhat":"Se vi trovate in aree con segnale debole o scarsa ricezione, la rete potrebbe non essere disponibile. I problemi di connettività potrebbero essere dovuti a problemi del gestore telefonico. Spostatevi in un'altra area e riprovate.^p","changeTo":""},
				{"findWhat":"Se doveste utilizzare il dispositivo mentre vi spostate, i servizi di rete potrebbero essere disabilitati a causa di problemi con la rete del gestore telefonico.^p","changeTo":""},
				{"findWhat":" dal gestore di rete o","changeTo":""},
			]

			for(var j=0; j < teXts.length; j++) {
				app.findTextPreferences.findWhat = teXts[j].findWhat;
				app.changeTextPreferences.changeTo = teXts[j].changeTo;
				doc[i].changeText();
				}
			
			app.findTextPreferences = NothingEnum.nothing;
			app.changeTextPreferences = NothingEnum.nothing;
		} 
		alert ("Service provider 문구 삭제 완료");
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_07.onClick = function () {
	var doc = app.documents;
	var status_icon_para = "Le icone e il loro aspetto potrebbero variare in base al gestore telefonico o al modello.";
	for (var i=0; i < doc.length; i++){
		//alert ("찾았다!");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = status_icon_para;
		var founds = doc[i].findText();
		var counter = 0
		for (var j=0; j < founds.length; j++) {
			app.activeDocument = doc[i];
			counter ++;
			} //alert (counter);
			if (counter > 0) {
				alert ("상태 아이콘 알림 문구가 있습니다.");
				founds[0].select();
				app.activeWindow.activePage = founds[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else if (counter == 0) {
				continue;
			}
		} 
		alert ("상태 아이콘 알림 문구가 없습니다. 알맞은 위치에 사양을 추가했으니 내용을 확인하세요.");
		main();
	function main() {
		//alert ("고고");
		var doc = app.documents;
		var find_para = "Alcune icone compaiono solo quando aprite il pannello delle notifiche.";
		for (var i=0; i < doc.length; i++){
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = find_para;
			var founds = doc[i].findText();
			for (var j=0; j < founds.length; j++) {
				app.activeDocument = doc[i];
				founds[j].select();
				app.activeWindow.activePage = founds[j].parentTextFrames[0].parentPage;
				founds[j].insertionPoints[-1].contents = "\r" + status_icon_para;
			}
		}
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_08.onClick = function () {
	var doc = app.documents;
	var dest_file = "005_Settings.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			//alert ("찾았다!");
			changeSettings();
			count ++;
		}
	}
	alert (dest_file + " 문서가 없습니다.");
	function changeSettings () {
		var myDoc = app.activeDocument;
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "Impostaz.";
		app.findTextPreferences.appliedParagraphStyle = "Chapter";
		app.changeTextPreferences.changeTo = "Impostazioni";
		var founds = myDoc.findText();
		for (var i=0; i<founds.length; i++) {
			founds[i].select();
			app.activeWindow.activePage = founds[i].parentTextFrames[0].parentPage;
		}
		myDoc.changeText();
		alert("Settings 챕터 번역 Full Text로 적용했습니다.");
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_09.onClick = function () {
	var doc = app.documents;
	var beforeText = "Se vi recate nel negozio in cui è stato effettuato il pagamento, è possibile effettuare lo storno appoggiando il vostro dispositivo al POS o lettore carte per finalizzare la restituzione dell'importo dell'acquisto.^pNell'elenco delle carte, scorrete verso sinistra o verso destra per selezionare la carta utilizzata. Seguite le istruzioni visualizzate per completare l'annullamento del pagamento.^pLa cancellazione di un pagamento effettuato con Samsung Pay segue le stesse regole di un pagamento effettuato con la carta fisica.^pPer maggiori informazioni, vi invitiamo a contattare la vostra banca o l'emittente della carta.";

	var changeText = "Potete annullare un pagamento recandovi presso il luogo dove è stato eseguito.^pNell'elenco delle carte, scorrete verso sinistra o verso destra per selezionare la carta utilizzata. Seguite le istruzioni visualizzate per completare l'annullamento del pagamento."
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = beforeText;
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("삼성페이 문구가 적용되어 있습니다.");
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				exit ();
			} else {
				app.findTextPreferences.findWhat = changeText;
				var finds = doc[i].findText();
				if (finds.length > 0) { // something has been found
					//alert("찾았다");
					app.activeDocument = doc[i];
					app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
					//finds[0].select();
					finds[0].select();
					app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
					main();
				} if (finds.length == 0) {
				}
				else {
					continue;
				}
			}
		} catch(e) {}
	}
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
	alert("삼성페이 문구를 추가할 위치를 찾을 수 없습니다.\n스크립트 수정이 필요하니 개발자에게 문의하시기 바랍니다.")
	exit();

	function main() {
		var doc = app.activeDocument;
		//템플릿 파일 열기
		var destination_doc = app.activeDocument;
		var source_file = File("C:/MA-WikiSpec/ITA/MA_online_UM_Ita_Samsung Pay_spec.indd"); // 사양문장 템플릿 문서 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Ita_Samsung Pay_spec.indd");
		var sourcePages = source_doc.pages[0];
		var frame = sourcePages.textFrames[0];
		var selection = frame.texts.everyItem().select();
		app.copy();
		source_doc.close();

		// app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		// app.findTextPreferences.findWhat = "######";
		// var foundSharp = app.activeDocument.findText();
		// foundSharp[0].select();

		destination_doc;
		app.paste();
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		alert("삼성페이 문구를 교체했습니다. 문서를 확인하세요.");
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------
w.show();