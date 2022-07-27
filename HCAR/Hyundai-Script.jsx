#targetengine "session";
#include "functions/rtolstylechanger.jsx";
#include "functions/batchexportidml.jsx"
#include "functions/update_path_links.jsx"
// #include "functions/htmlforxml.jsx"
#include "functions/translateforxml.jsx"
#include "functions/Package_book_idml.jsx"
#include "functions/Package_docu_idml.jsx"
#include "functions/batcheXportpdf.jsx"
#include "functions/updatetoc.jsx"
#include "functions/titlechangeCase.jsx"
#include "functions/checkInlineimage.jsx"
#include "functions/findNot100img.jsx"
#include "functions/fixPositionTextframe.jsx"
#include "functions/notrans2pdf.jsx"
#include "functions/eXportBook2pdf.jsx"
#include "functions/revision.jsx"

var script = app.activeScript;
var myScriptFolderPath = script.path + '/functions';
var noDocuType = "설정된 문서 범위가 없습니다.";
var noDoc = "현재 열려있는 문서가 없습니다.";

var w = new Window ("palette","MC:Hyundai Script ver 0.0.9"); // 22.01.03 업데이트
	var group = w.add ('group {orientation: "column"}');
	var panel_00 = group.add ('panel {text: "RtoL 스타일 변환 옵션"}');
	var btn_01 = panel_00.add("button", [0, 0, 200, 25], "RtoL 단락/문자 스타일 변환");
	var btn_02 = panel_00.add("button", [0, 0, 200, 25], "바인딩/스토리/테이블 방향 전환");
	var btn_03 = panel_00.add("button", [0, 0, 200, 25], "선택한 테이블 RtoL 전환");
	var btn_04 = panel_00.add("button", [0, 0, 200, 25], "선택한 테이블 LtoR 전환");

	var panel_02 = group.add ('panel {text: "내보내기 옵션"}');
	var btn_13 = panel_02.add("button", [0, 0, 200, 25], "파일 별 PDF 내보내기");
	var btn_21 = panel_02.add("button", [0, 0, 200, 25], "북파일 PDF 내보내기");
	var btn_07 = panel_02.add("button", [0, 0, 200, 25], "IDML 파일 내보내기");
	// var btn_09 = panel_02.add("button", [0, 0, 200, 25], "HTML 변환용 XML 내보내기");
	// var btn_10 = panel_02.add("button", [0, 0, 200, 25], "번역용 XML 내보내기");
	var btn_11 = panel_02.add("button", [0, 0, 200, 25], "북파일 패키지(with IDML)");
	var btn_12 = panel_02.add("button", [0, 0, 200, 25], "단일문서 패키지(with IDML)");
	var btn_22 = panel_02.add("button", [0, 0, 200, 25], "리비전 파일 만들기");
	// btn_16 = panel_02.add("button", [0, 0, 200, 25], "북 페이지 이동");

	var panel_04 = group.add ('panel {text: "문서 편집 옵션1"}');
	var btn_16 = panel_04.add("button", [0, 0, 200, 25], "인덱스 업데이트");
	var btn_08 = panel_04.add("button", [0, 0, 200, 25], "이미지 경로 재설정");
	var btn_18 = panel_04.add("button", [0, 0, 200, 25], "100%가 아닌 아이콘 찾기");
	
	var panel_03 = group.add ('panel {text: "문서 편집 옵션2"}');
	var btn_14 = panel_03.add("button", [0, 0, 200, 25], "목차 업데이트");
	var btn_15 = panel_03.add("button", [0, 0, 200, 25], "타이틀 첫글자 대문자 변경");
	var btn_17 = panel_03.add("button", [0, 0, 200, 25], "100%가 아닌 이미지 목록 추출");
	var btn_19 = panel_03.add("button", [0, 0, 200, 25], "텍스트 프레임 위치 자동 정렬");
	var btn_20 = panel_03.add("button", [0, 0, 200, 25], "번역 PDF 출력하기");
	var radio01 = panel_03.add("radiobutton", [0, 0, 180, 20], "현재 열려 있는 문서");
	var radio02 = panel_03.add("radiobutton", [0, 0, 180, 20], "열려 있는 모든 문서");

w.show();

function selectedDocuType (radios) {
	for (var i=0; i<radios.children.length; i++) {
		if (radios.children[i].value == true) {
			if (radios.children[i].text == "현재 열려 있는 문서") {
				return "activeDocument";
			} else if (radios.children[i].text == "열려 있는 모든 문서") {
				return "documents"
			}
		}
	}
}

btn_01.onClick = function() {
	if (app.documents.length == 0) {
		alert("인디자인 문서를 열고 실행하세요.");
		exit();
	}
	else if (app.documents.length == 1) {
		w.close();
		var myWindow = new Window ('palette', "작업 진행 중 ...");
			myWindow.pbar = myWindow.add ('progressbar', undefined, 0, 4);
			myWindow.pbar.preferredSize.width = 300;
		myWindow.show();
		
		RtoL_01 ();
		myWindow.pbar.value = 1;
		RtoL_03 ();
		myWindow.pbar.value = 2;
		RtoL_02 ();
		myWindow.pbar.value = 3;
		RtoL_04 (); //arabic digit
		myWindow.pbar.value = 4;
		
		myWindow.close();
		alert ("LtoR을 RtoL로 변환 완료");
		w.show();
	}
	else if (app.documents.length > 1) {
		var docs = app.documents;
		w.close();
		var myWindow = new Window ('palette', "작업 진행 중 ...");
			myWindow.pbar = myWindow.add ('progressbar', undefined, 0, docs.length);
			myWindow.pbar.preferredSize.width = 300;
		myWindow.show();
		
		for (j=0; j<docs.length; j++) {
			myWindow.pbar.value = j+1;
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			RtoL_01 ();
			RtoL_03 ();
			RtoL_02 ();
			RtoL_04 ();
			app.activeDocument.save();
		}
		myWindow.close();
		alert(docs.length + " 개의 문서, LtoR을 RtoL로 변환 완료");
		w.show();
	}
}

btn_02.onClick = function() {
	if (app.documents.length == 0) {
		alert("인디자인 문서를 열고 실행하세요.");
		exit();
	}
	else if (app.documents.length == 1) {
		directionTransform();
		moveTextFrame();
		alert ("바인딩, 스토리, 테이블의 방향 변경 완료")
	}
	else if (app.documents.length > 1) {
		var docs = app.documents;
		w.close();
		var myWindow = new Window ('palette', "작업 진행 중 ...");
			myWindow.pbar = myWindow.add ('progressbar', undefined, 0, docs.length);
			myWindow.pbar.preferredSize.width = 300;
		myWindow.show();
		
		for (j=0; j<docs.length; j++) {
			myWindow.pbar.value = j+1;
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			directionTransform();
			moveTextFrame();
			app.activeDocument.save();
		}
		myWindow.close();
		alert(docs.length + " 개의 문서, 바인딩, 스토리, 테이블의 방향 변경 완료");
		w.show();
	}
}

btn_03.onClick = function() {
	var selTable = app.selection[0];
	selTable.tableDirection = TableDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
}

btn_04.onClick = function() {
	var selTable = app.selection[0];
    selTable.tableDirection = TableDirectionOptions.LEFT_TO_RIGHT_DIRECTION;
}

// btn_05.onClick = function() {
// 	fixedObjectXoffset();
// }

// btn_06.onClick = function() {
// 	if (app.documents.length == 0) {
// 		alert("인디자인 문서를 열고 실행하세요.");
// 		exit();
// 	}
// 	List_pscs();
// }

btn_07.onClick = function() {
	// alert(app.documents.length);
	if (app.documents.length == 1) {
		var doc = app.activeDocument;
		myFilePath = doc.filePath.absoluteURI + "/" + doc.name.split(".indd")[0] + ".idml"; //파일명 설정
		myFile = new File(myFilePath);
		doc.exportFile(ExportFormat.INDESIGN_MARKUP, myFile); 
		alert("IDML 출력 완료");
	} else if (app.documents.length == 0) {
		w.close();
		batchExportIDML();
		w.show();
	}
}

btn_08.onClick = function() {
	w.close();
	CreateDialog();
	w.show();
}

// btn_09.onClick = function() {
// 	w.close();
// 	htmlforXML();
// 	w.show();
// }

// btn_10.onClick = function() {
// 	w.close();
// 	translateforXML(myScriptFolderPath);
// 	w.show();
// }

btn_11.onClick = function() {
	w.close();
	PreCheck();
	w.show();
}

btn_12.onClick = function() {
	w.close();
	PreCheckDocu();
	w.show();
}

btn_13.onClick = function() {
	if (app.books.length == 0) {
		alert("북 파일이 열려 있지 않습니다.");
		exit();
	} else if (app.books.length > 1) {
		alert("하나의 북 파일만 실행하세요.");
		exit();
	} else {
		w.close();
		batcheXportpdf();
		w.show();
	}
}

btn_14.onClick = function() {
	if (app.documents.length == 0) {
		alert(noDoc);
		exit();
	}
	
	docuType = selectedDocuType (panel_03);
	if (docuType == null) {
		alert(noDocuType);
		exit();
	} else if (docuType == "activeDocument") {
		updateAllToc(0);
	} else if (docuType == "documents") {
		updateAllToc(1);
	}
}

btn_15.onClick = function() {
	if (app.documents.length == 0) {
		alert(noDoc);
		exit();
	}
	w.close();
	docuType = selectedDocuType (panel_03);
	if (docuType == null) {
		alert(noDocuType);
		w.show();
	} else if (docuType == "activeDocument") {
		var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, 1);
			myProg.pbar.preferredSize.width = 300;
			myProg.show();
			myProg.pbar.value = 1;
		titleChangeCaseUpper();
		app.activeDocument.save();
		myProg.close();
		alert("완료합니다.");
		w.show();
	} else if (docuType == "documents") {
		var docs = app.documents;
		var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, docs.length);
			myProg.pbar.preferredSize.width = 300;
			myProg.show();
		for (var n=0; n<docs.length; n++) {
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			myProg.text = app.activeDocument.name + ' 업데이트 중...';
			myProg.pbar.value = n + 1;	
			titleChangeCaseUpper()
			app.activeDocument.save();
		}
		myProg.close();
		alert("완료합니다.");
		w.show();
	}
}

btn_17.onClick = function() {
	if (app.documents.length == 0) {
		alert(noDoc);
		exit();
	}
	w.close();
	docuType = selectedDocuType (panel_03);
	if (docuType == null) {
		alert(noDocuType);
		w.show();
	} else if (docuType == "activeDocument") {
		var myDoc = app.activeDocument;
		CheckImageSize(myDoc);
		alert("완료합니다.");
		w.show();
	} else if (docuType == "documents") {
		var docs = app.documents;
		var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, docs.length);
			myProg.pbar.preferredSize.width = 300;
			myProg.show();
		for (var n=0; n<docs.length; n++) {
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			myProg.text = app.activeDocument.name + ' 업데이트 중...';
			myProg.pbar.value = n + 1;	
			var myDoc = app.activeDocument;
			CheckImageSize(myDoc);
		}
		myProg.close();
		alert("완료합니다.");
		w.show();
	}
}

btn_19.onClick = function() {
	if (app.documents.length == 0) {
		alert(noDoc);
		exit();
	}
	w.close();
	docuType = selectedDocuType (panel_03);
	if (docuType == null) {
		alert(noDocuType);
		w.show();
	} else if (docuType == "activeDocument") {
		var myDoc = app.activeDocument;
		fixTextFrames(myDoc);
		alert("완료합니다.");
		w.show();
	} else if (docuType == "documents") {
		var docs = app.documents;
		var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, docs.length);
			myProg.pbar.preferredSize.width = 300;
			myProg.show();
		for (var n=0; n<docs.length; n++) {
			var openDocs = app.documents.everyItem().getElements();
			app.activeDocument = openDocs[openDocs.length-1];
			myProg.text = app.activeDocument.name + ' 업데이트 중...';
			myProg.pbar.value = n + 1;	
			var myDoc = app.activeDocument;
			fixTextFrames(myDoc);
		}
		myProg.close();
		alert("완료합니다.");
		w.show();
	}
}

btn_16.onClick = function() { // 인덱스 업데이트
	// 인덱스 문서가 열려있는 경우
	// 문서가 열려있지 않다면 다른 모든 문서를 닫을 것을 경고
	// 북파일이 열려있는지 확인
	
	if (app.documents.length > 0) {
		alert("열려 있는 모든 문서를 닫은 후 실행하세요.");
		exit();
	}
	if (app.books.length == 0) {
		alert("현재 열려 있는 북 파일이 없습니다.");
		exit();
	} else {
		w.close();
		updateIndex();
		w.show();
	}
}

btn_18.onClick = function() { // 100%가 아닌 아이콘 이미지 찾기
	if (app.documents.length == 0) {
		alert("현재 열려 있는 문서가 없습니다.");
		exit();
	} else {
		w.close();
		checkInlineImages();
		// w.show();
	}
}

btn_20.onClick = function() { // 번역 PDF 출력하기
	if (app.documents.length == 0) {
		alert(noDoc);
		exit();
	}
	w.close();
	docuType = selectedDocuType (panel_03);
	if (docuType == null) {
		alert(noDocuType);
		w.show();
	} else if (docuType == "activeDocument") {
		var exportPath = Folder.selectDialog("PDF를 출력할 폴더를 선택하세요." );
		if (exportPath == null) {
			exit();
		}
		var myDoc = app.activeDocument;
		_notrans2pdf(myDoc, exportPath);
		alert("완료합니다.");
		w.show();
	} else if (docuType == "documents") {
		var exportPath = Folder.selectDialog("PDF를 출력할 폴더를 선택하세요." );
		if (exportPath == null) {
			exit();
		}
		var docs = app.documents;
		var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, docs.length);
			myProg.pbar.preferredSize.width = 300;
			myProg.show();
		for (var n=0; n<docs.length; n++) {
			_notrans2pdf(docs[n], exportPath);
		}
		myProg.close();
		alert("완료합니다.");
		w.show();
	}
}

btn_21.onClick = function() {
	if (app.books.length > 0 || app.documents.length > 0) {
		alert("개별 문서 또는 북 파일을 모두 닫은 후에 다시 실행하세요.");
	} else {
		w.close();
		_eXportBook2pdf();
		alert("완료합니다.");
		w.show();
	}
}

btn_22.onClick = function() {
	if (app.books.length == 0) {
		alert("열려있는 북 파일이 없습니다.");
	} else {
		w.close();
		_reVision();
		alert("완료합니다.");
		w.show();
	}
}