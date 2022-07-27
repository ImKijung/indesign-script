#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";
#include "_reference/changeCapDown.jsx";

var w = new Window ('palette {text: "LIT 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "콜론 뒤 소문자 변경");
		btn_04 = group02.add("button", [0, 0, 150, 30], "헤딩 내 따옴표 확인");

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
			app.findTextPreferences.findWhat = "Akumuliatoriaus išėmimas";
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
	headingMMIXport()
}
//----------------------------------------------------------------------------------------------------
w.show();

function headingMMIXport() {
	var docs = app.documents;
	if (docs.length == 0) {
		alert ("열려있는 인디자인 문서가 없습니다.");
	} else {
		var fdoc = app.activeDocument;
		var docPath = fdoc.filePath;
		var file = new File(docPath + "/LIT-Heading-list.txt");
		file.open("w");
		file.encoding = "utf-8";
		var myWindow = new Window ('palette');
			myWindow.pbar = myWindow.add ('progressbar', undefined, 0, docs.length);
			myWindow.pbar.preferredSize.width = 150;
			myWindow.show();
		for (k=0;k<docs.length; k++) {
			myWindow.pbar.value = k+1;
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			var doc = app.activeDocument;
			var story = doc.stories[0];
			var paras = story.paragraphs.everyItem().getElements();
			for (var i=0;i<paras.length;i++) {
				var para = paras[i];
				var pNum = para.parentTextFrames[0].parentPage.name;
				if (para.appliedParagraphStyle.name.indexOf("Heading1") != -1) {
					var tsr = para.textStyleRanges;
					for (var j=0;j<tsr.length;j++) {
						if (tsr[j].appliedCharacterStyle.name.indexOf("MMI") != -1) {
							// var pContents = para.contents.replace("\x{FEFF}","");//\x{FEFF}
							// $.writeln(doc.name + " - " + pNum + " - " + para.appliedParagraphStyle.name + " - " + para.contents);
							file.writeln(doc.name + " - " + pNum + " - " + para.appliedParagraphStyle.name + " - " + para.contents);
							// $.writeln(tsr[j].contents)
						}
					}
				}
			}
			$.sleep(20); // Do something useful here
		}
		alert("LIT 헤딩 MMI 목록 추출 완료");
	}
	file.close();
	file.execute();
}