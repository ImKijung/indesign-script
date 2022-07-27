#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapUp.jsx";

var w = new Window ('palette {text: "TUR 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_03 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 삭제");
	group01 = w.add ('group {orientation: "column"}');
	group01.orientation = "row";
		btn_02 = group01.add("button", [0, 0, 150, 30], "팻네임 넣기");
		btn_04 = group01.add("button", [0, 0, 150, 30], "기기 종류 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_05 = group02.add("button", [0, 0, 150, 30], "Wi-Fi 주의 문구 적용");
		btn_06 = group02.add("button", [0, 0, 150, 30], "Wi-Fi 지원 국가 목록 추가");
	group03 = w.add ('group {orientation: "column"}');
	group03.orientation = "row";
		btn_07 = group03.add("button", [0, 0, 150, 30], "제조사/법인/보증기간/콜센터");
		btn_08 = group03.add("button", [0, 0, 150, 30], "배터리 분리 어려움 추가");
	group04 = w.add ('group {orientation: "column"}');
	group04.orientation = "row";
		btn_09 = group04.add("button", [0, 0, 150, 30], "3GPP SMS 문구 추가");
		btn_10 = group04.add("button", [0, 0, 150, 30], "터키 규정 준수 문구 확인");
	group05 = w.add ('group {orientation: "column"}');
	group05.orientation = "row";
		btn_11 = group05.add("button", [0, 0, 150, 30], "4.5G 아이콘 추가");
		btn_12 = group05.add("button", [0, 0, 150, 30], "번인 문구 확인");

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
			app.findTextPreferences.findWhat = "Pilin çıkartılması";
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
btn_02.onClick = function () {
	var doc = app.documents;
	var dest_file = "001_Cover.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			alert ("팻네임이 적용됐는지 문서를 확인하세요.");
		}
	}
}
//----------------------------------------------------------------------------------------------------
btn_04.onClick = function () {
	var doc = app.documents;
	var dest_file = "001_Cover.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			alert ("기기 종류가 User Manual 하단에 적용되어 있는지 확인하세요. 다음 팝업 창에서 종류에 맞는 기기의 내용을 복사해 입력하세요.");
		} else {}
	}
	var myWindow = new Window ("dialog", "기기 종류");
		myWindow.alignment = "left";
		var myInput00 = myWindow.add("group");
			myInput00.add("edittext",[0,0,270,70], "전화 기능이 있어야만 Mobile Phone으로 구분\r\r태블릿의 경우, 전화 기능이 있다면 Mobile Phone으로 간주",{multiline: true,readonly: true})
		var myInput01 = myWindow.add("group");
			myInput01.add("statictext",[0,0,80,20],"Mobile Phone");
			var myText01 = myInput01.add("edittext",[0,0,180,20],"Cep Telefonu",{readonly: true});
				myText01.characters = 20;
				myText01.active = true;
		var myInput02 = myWindow.add("group");
			myInput02.add("statictext",[0,0,80,20],"Tablet");
			var myText02 = myInput02.add("edittext",[0,0,180,20],"Mobil İnternet Cihazı",{readonly: true});
		var myInput03 = myWindow.add("group");
			myInput03.add("statictext",[0,0,80,20],"Camera");
			var myText03 = myInput03.add("edittext",[0,0,180,20],"Kamera",{readonly: true});
		var myInput04 = myWindow.add("group");
			myInput04.add("statictext",[0,0,80,20],"HomeSync");
			var myText04 = myInput04.add("edittext",[0,0,180,20],"Homesync Kişisel Bilgisayar",{readonly: true});
	myWindow.show ();
}
//----------------------------------------------------------------------------------------------------
btn_05.onClick = function () {
	alert("Wi-Fi 2.4+5GHz 또는 5GHz 지원 모델인지 확인하세요. 2.4GHz만 지원하는 경우 주의 문구가 적용되지 않도록 주의하세요.\rOnline UM에서는 사양을 확인할 수 없으니 PM에게 문의 후 진행하세요.")
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "WLAN bandının kullanımı yalnızca kapalı alanda kullanımla sınırlıdır. Bu kısıtlama aşağıdaki tüm ülkelerde uygulanacaktır.";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("Wi-Fi 주의 문구가 이미 적용되어 있습니다.");
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
		app.findTextPreferences.findWhat = "Wi-Fi özelliğini etkinleştirerek bir Wi-Fi ağına bağlanın; İnternet ve diğer ağ cihazlarına erişim sağlayın.";
		app.findTextPreferences.appliedParagraphStyle = "Description";
		var founds = doc[i].findText();
		if (founds.length > 0) {
			app.activeDocument = doc[i];
			founds[0].select();
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
		var source_file = File("C:/MA-WikiSpec/TUR/MA_online_UM_Tur_Wi-Fi_Notice_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Tur_Wi-Fi_Notice_spec.indd");
		var sourcePages = source_doc.pages[0];
		var frame = sourcePages.textFrames[0];
		var selection = frame.texts.everyItem().select();
		app.copy();
		source_doc.close();
		
		destination_doc;
		alert("Wi-Fi 주의 문구를 클립보드에 복사했습니다. 적용해야할 위치에 붙여넣기하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_06.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++) {
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		app.findTextPreferences.findWhat = "Wi-Fi özelliğini etkinleştirerek bir Wi-Fi ağına bağlanın; İnternet ve diğer ağ cihazlarına erişim sağlayın.";
		app.findTextPreferences.appliedParagraphStyle = "Description";
		var founds = doc[i].findText();
		if (founds.length > 0) {
			app.activeDocument = doc[i];
			founds[0].select();
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
		var source_file = File("C:/MA-WikiSpec/TUR/MA_online_UM_Tur_Wi-Fi_Countries_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Tur_Wi-Fi_Countries_spec.indd");
		var sourcePages = source_doc.pages[0];
		var frame = sourcePages.textFrames[0];
		var selection = frame.texts.everyItem().select();
		app.copy();
		source_doc.close();
		
		destination_doc;
		alert("Wi-Fi 사용 국가리스트 템플릿을 클립보드에 복사했습니다. 적용해야할 위치에 붙여넣기하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_07.onClick = function () {
	var doc = app.documents;
	var dest_file = "007_Copyright.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Cihazın kullanım ömrü 5 yıl, garanti süresi 2 yıldır.";
				var founds00 = app.activeDocument.findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("제조사/법인/보증기간/콜센터가 이미 적용되어 있습니다.");
					exit ();
				} else {}
			} catch(e) {}

			//alert ("찾았다!");
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Telif hakkı";
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
		var source_file = File("C:/MA-WikiSpec/TUR/MA_online_UM_Tur_Copyright_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Tur_Copyright_spec.indd");
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
		alert("제조사/법인/보증기간/콜센터를 추가했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_08.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Bu üründeki batarya(lar) kullanıcılar tarafından kolaylıkla değiştirilemez.";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("배터리 분리 어려움 문구가 이미 적용되어 있습니다.");
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
		app.findTextPreferences.findWhat = "Enerjiden tasarruf etmek için kullanmadığınızda şarj cihazını elektrik prizinden çıkarın. Şarj cihazında güç anahtarı yoktur, dolayısıyla elektriği boşa harcamamak için kullanmadığınızda şarj cihazını elektrik prizinden çıkarmanız gereklidir. Şarj cihazı elektrik prizine yakın durmalı ve şarj sırasında kolay erişilmelidir.";
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

	function main() {
		var destination_doc = app.activeDocument;
		//템플릿 파일 열기
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var source_file = File("C:/MA-WikiSpec/TUR/MA_online_UM_Tur_Battery_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Tur_Battery_spec.indd");
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
		alert("배터리 분리 어려움 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_09.onClick = function () {
	var doc = app.documents;
	var countLTE = 0;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Mesajlar";
			app.findTextPreferences.appliedParagraphStyle = "Heading1-APPLINK"
			var foundLTE = doc[i].findText();
			for (var j=0; j < foundLTE.length; j++) {
				countLTE ++;
				if (countLTE > 0) {
					alert ("메시지 앱이 매뉴얼에 포함되어 있습니다. 3GPP SMS 문구를 추가합니다.");
					gogo();
				} else if (countLTE < 1) {
					alert ("메시지 앱이 매뉴얼에 포함되어 있지 않습니다.");
					exit();
				} else {}
			}
		} catch(e) {}
	}
	function gogo() {
		for (var i=0; i < doc.length; i++){
			try {
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				app.findTextPreferences.findWhat = "Bu cihaz Türkçe karakterlerin tamamını ihtiva eden ETSI TS 123.038 V8.0.0 ve ETSI TS 123.040 V8.1.0 teknik özelliklerine uygundur.";
				var founds00 = doc[i].findText();
				var counter00 = 0;
				for (var j=0; j<founds00.length; j++) {
					counter00 ++;
				} if (counter00 > 0) {
					alert ("3GPP SMS 문구가 이미 적용되어 있습니다.");
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
			app.findTextPreferences.findWhat = "Dolaşımda iken mesaj göndermek için ek ücret ödemeniz gerekebilir.";
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
		var source_file = File("C:/MA-WikiSpec/TUR/MA_online_UM_Tur_3GPP_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Tur_3GPP_spec.indd");
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
		alert("3GPP SMS 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_10.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Bu cihaz Türkiye altyapısına uygundur.";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("터키 규정 준수 문구가 이미 적용되어 있습니다.");
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
		app.findTextPreferences.findWhat = "Talimat simgeleri";
		app.findTextPreferences.appliedParagraphStyle = "Heading3"
		var founds = doc[i].findText();
		if (founds.length > 0) {
			app.activeDocument = doc[i];
			founds[0].insertionPoints[0].contents= "######\r";
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
		var source_file = File("C:/MA-WikiSpec/TUR/MA_online_UM_Tur_Regulations_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Tur_Regulations_spec.indd");
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
		alert("터키 규정 준수 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
btn_11.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findChangeTextOptions.caseSensitive = true;
			app.findTextPreferences.findWhat = "LTE şebekesi bağlı";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("4.5G 아이콘이 적용됐는지 확인하세요. 아이콘이 없다면 수동으로 추가해주세요.");
				app.activeDocument = doc[i];
				founds00[0].select();
				app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			} else {}
		} catch(e) {}
	}
}
//----------------------------------------------------------------------------------------------------
btn_12.onClick = function () {
	var doc = app.documents;
	for (var i=0; i < doc.length; i++){
		try {
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			app.findTextPreferences.findWhat = "Dokunmatik ekranın bir bölümünde veya tamamında uzun süre sabit görüntü bırakmamanız önerilir. Bu durum kalıntı görüntü (ekran yanması) veya gölge görüntü oluşmasına sebep olabilir.";
			var founds00 = doc[i].findText();
			var counter00 = 0;
			for (var j=0; j<founds00.length; j++) {
				counter00 ++;
			} if (counter00 > 0) {
				alert ("번인 문구가 이미 적용되어 있습니다.");
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
		app.findTextPreferences.findWhat = "Dokunmatik ekrana zarar vermemek için keskin subjeler ile dokunmayın veya parmak uçlarınız ile aşırı basınç uygulamayın.";
		var founds = doc[i].findText();
		if (founds.length > 0) {
			app.activeDocument = doc[i];
			app.activeWindow.activePage = founds00[0].parentTextFrames[0].parentPage;
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
		var source_file = File("C:/MA-WikiSpec/TUR/MA_online_UM_Tur_Burn-in_spec.indd"); // 사양문장 템플릿 문서 열기
		app.open(source_file);
		var source_doc = app.documents.item("MA_online_UM_Tur_Burn-in_spec.indd");
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
		alert("번인 문구를 적용했습니다. 문서를 확인하세요.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
w.show();