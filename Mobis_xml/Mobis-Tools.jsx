#targetengine "session";
#include "Functions/commonLib.jsx";
#include "Functions/batchExportIDML.jsx";
#include "Functions/setImagePath.jsx";
#include "Functions/hidelayers.jsx";
#include "Functions/visiblelayers.jsx";
#include "Functions/addLayers.jsx";
#include "Functions/deletelayers.jsx";
#include "Functions/eXportBook.jsx";
#include "Functions/updatetoc.jsx";
#include "Functions/_eXporthtml-img.jsx";
#include "Functions/find-Empty.jsx";

var main = new Window ("palette", "Mobis Tools v.0.1.4");
var panel_02 = main.add ('panel {text: "Only Book"}');
var btn_01 = panel_02.add("button", [0, 0, 180, 30], "IDML 파일 내보내기");
var btn_07 = panel_02.add("button", [0, 0, 180, 30], "북 파일 순서 확인");
var btn_10 = panel_02.add("button", [0, 0, 180, 30], "목차/색인 업데이트");
var btn_08 = panel_02.add("button", [0, 0, 180, 30], "HTML 이미지 만들기");

var panel_03 = main.add ('panel {text: "PDF"}');
var btn_11 = panel_03.add("button", [0, 0, 180, 30], "검토용 PDF");
var btn_12 = panel_03.add("button", [0, 0, 180, 30], "최종 검토용 PDF");
var btn_13 = panel_03.add("button", [0, 0, 180, 30], "보안용 PDF");
var btn_14 = panel_03.add("button", [0, 0, 180, 30], "인쇄용 PDF (4도)");
var btn_15 = panel_03.add("button", [0, 0, 180, 30], "인쇄용 PDF (1도)");

var panel_01 = main.add ('panel {text: "Options"}');
var btn_02 = panel_01.add("button", [0, 0, 180, 30], "전체 이미지 링크 재설정");
var btn_03 = panel_01.add("button", [0, 0, 180, 30], "레이어 숨기기");
var btn_04 = panel_01.add("button", [0, 0, 180, 30], "레이어 보이기");
var btn_05 = panel_01.add("button", [0, 0, 180, 30], "레이어 추가");
var btn_06 = panel_01.add("button", [0, 0, 180, 30], "레이어 삭제");
var btn_16 = panel_01.add("button", [0, 0, 180, 30], "Empty 표 검증");


panel_02 = main.add ('panel {text: "문서 범위"}');
panel_02.alignChildren = "left";
var radio01 = panel_02.add("radiobutton", [0, 0, 180, 20], "현재 열려 있는 문서");
var radio02 = panel_02.add("radiobutton", [0, 0, 180, 20], "열려 있는 모든 문서");
var radio03 = panel_02.add("radiobutton", [0, 0, 180, 20], "활성화되어 있는 북 파일");

var noDoc = "현재 열려있는 문서가 없습니다.";
var noDocuType = "설정된 문서 범위가 없습니다.";
var noBook = "현재 활성화된 북 파일이 없습니다.";

btn_01.onClick = function() { //IDML 파일 내보내기
	main.close();
	batchExportIDML();
	main.show();
}

btn_02.onClick = function() { //전체 이미지 링크 재설정
	docuType = selectedDocuType (panel_02);
	if (docuType == null) {
		alert(noDocuType);
	} else if (docuType == "activeDocument") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			rePlaceImagePath(docuType);
			main.show();
		}
	} else if (docuType == "documents") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			rePlaceImagePath(docuType);
			main.show();
		}
	} else if (docuType == "activeBook") {
		if (app.books.length == 0) {
			alert(noBook);
		} else if (app.books.length > 0) {
			main.close();
			rePlaceImagePath(docuType);
			main.show();
		}
	}
}

btn_03.onClick = function() { //레이어 숨기기
	docuType = selectedDocuType (panel_02);
	if (docuType == null) {
		alert(noDocuType);
	} else if (docuType == "activeDocument") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			HideLayers(docuType);
			main.show();
		}
	} else if (docuType == "documents") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			HideLayers(docuType);
			main.show();
		}
	} else if (docuType == "activeBook") {
		if (app.books.length == 0) {
			alert(noBook);
		} else if (app.books.length > 0) {
			main.close();
			HideLayers(docuType);
			main.show();
		}
	}
}

btn_04.onClick = function() { //레이어 보이기
	docuType = selectedDocuType (panel_02);
	if (docuType == null) {
		alert(noDocuType);
	} else if (docuType == "activeDocument") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			VisibleLayers(docuType);
			main.show();
		}
	} else if (docuType == "documents") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			VisibleLayers(docuType);
			main.show();
		}
	} else if (docuType == "activeBook") {
		if (app.books.length == 0) {
			alert(noBook);
		} else if (app.books.length > 0) {
			main.close();
			VisibleLayers(docuType);
			main.show();
		}
	}
}

btn_05.onClick = function() { //레이어 추가
	docuType = selectedDocuType (panel_02);
	if (docuType == null) {
		alert(noDocuType);
	} else if (docuType == "activeDocument") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			AddLayer(docuType);
			main.show();
		}
	} else if (docuType == "documents") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			AddLayer(docuType);
			main.show();
		}
	} else if (docuType == "activeBook") {
		if (app.books.length == 0) {
			alert(noBook);
		} else if (app.books.length > 0) {
			main.close();
			AddLayer(docuType);
			main.show();
		}
	}
}

btn_06.onClick = function() { //레이어 삭제1111111
	docuType = selectedDocuType (panel_02);
	if (docuType == null) {
		alert(noDocuType);
	} else if (docuType == "activeDocument") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			DeleteLayers(docuType);
			main.show();
		}
	} else if (docuType == "documents") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			DeleteLayers(docuType);
			main.show();
		}
	} else if (docuType == "activeBook") {
		if (app.books.length == 0) {
			alert(noBook);
		} else if (app.books.length > 0) {
			main.close();
			DeleteLayers(docuType);
			main.show();
		}
	}
}

btn_16.onClick = function() {
	docuType = selectedDocuType (panel_02);
	if (docuType == null) {
		alert(noDocuType);
	} else if (docuType == "activeDocument") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			_findEmptyTalbe(docuType);
			main.show();
		}
	} else if (docuType == "documents") {
		if (app.documents.length == 0) {
			alert(noDoc);
		} else if (app.documents.length > 0) {
			main.close();
			_findEmptyTalbe(docuType);
			main.show();
		}
	} else if (docuType == "activeBook") {
		if (app.books.length == 0) {
			alert(noBook);
		} else if (app.books.length > 0) {
			main.close();
			_findEmptyTalbe(docuType);
			main.show();
		}
	}
}

btn_07.onClick = function() {
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		// main.close();
		checkBookFileOrder();
		// main.show();
	}
}

btn_08.onClick = function() {
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		main.close();
		var mode = eXportHTMLimg();
		if (mode == false) {
			alert("이미지 경로가 설정되어 있지 않거나 업데이트하지 않았습니다. 이미지 상태를 확인하세요.");
		}
		main.show();
	}
}

btn_11.onClick = function() { //검토용 PDF
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		var mode = "selectview";
		var preset = false;
		var pageType = true;//스프레드형식
		main.close();
		eXportBookConfirm(mode, preset, pageType, 0);
		main.show();
	}
}

btn_12.onClick = function() { //최종 검토용 PDF
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		var mode = "allview";
		var preset = false;
		var pageType = true;//스프레드형식
		main.close();
		eXportBookConfirm(mode, preset, pageType, 0);
		main.show();
	}
}

btn_13.onClick = function() { //보안 검토용 PDF
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		var mode = "none";
		var preset = false;
		var pageType = false;//페이지
		main.close();
		eXportBookConfirm(mode, preset, pageType, 0);
		main.show();
	}
}

btn_14.onClick = function() { //인쇄용 4도
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		var mode = "none";
		var preset = true; // 인쇄용
		var pageType = false; //페이지
		main.close();
		eXportBookConfirm(mode, preset, pageType, 0);
		main.show();
	}
}

btn_15.onClick = function() { //인쇄용 1도
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		var mode = "none";
		var preset = true; // 인쇄용
		var pageType = false; //페이지
		main.close();
		eXportBookConfirm(mode, preset, pageType, 1);
		main.show();
	}
}

btn_10.onClick = function() {
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.books.length > 0) {
		main.close();
		updateAllTocIndex();
		main.show();
	}
}

main.show();

function selectedDocuType (radios) {
	for (var i=0; i<radios.children.length; i++) {
		if (radios.children[i].value == true) {
			if (radios.children[i].text == "현재 열려 있는 문서") {
				return "activeDocument";
			} else if (radios.children[i].text == "열려 있는 모든 문서") {
				return "documents"
			} else if (radios.children[i].text == "활성화되어 있는 북 파일") {
				return "activeBook"
			}
		}
	}
}