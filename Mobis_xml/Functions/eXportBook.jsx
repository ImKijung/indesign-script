#include "commonLib.jsx";

// eXportBookConfirm("selectview", false, true, 0);
// 실행하기 전 확인하기
function eXportBookConfirm(mode, preset, pageType, colorMode) {
	var layerMode = mode;
	var pdfSet = preset;
	// 모든 문서가 열려있는지 확인한다.
	if (app.documents.length == 0) {
		alert("모든 문서를 열어놓고 실행하세요.");
		exit();
	} else if (app.books.length == 0) {
		alert("열려있는 북 파일이 없습니다.");
		exit();
	} else if (app.documents.length > 0) {
		var docuNum = app.documents.length;
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		var bookNum = myBookContents.length;
		if (docuNum != bookNum) {
			alert("북파일의 문서 개수와 열려있는 문서의 개수가 다릅니다.\r북 파일의 모든 문서를 열어 놓고 실행하세요.");
			exit();
		}
	} else {}
	
	
	var myDoc = app.documents;
	for (var lx=0; lx<myDoc.length; lx++) {
		if (myDoc[lx].name.indexOf("Cover") != -1) {
			continue
		} else {
			checkBlankPages(myDoc[lx]);
		}
		_checkImages(myDoc[lx]);
		_checkCross(myDoc[lx]);
	}

	for (var b=0; b<myBookContents.length; b++) {
		var myPath = myBookContents[b].fullName;
		var myDoc = File(myPath);
		app.open(myDoc);
		
	}
	
	if (mode == "allview") {
		// alert("레이어 전체 출력 모드");
		visibleAllLayers(docuNum);
	} else if (mode == "selectview") {
		// alert("차수별 레이어 출력 모드");
		visibleSelectedLayers(docuNum);
	} else if (mode == "none") {
		hideLayersAll();
	}

	var pdfPreset;

	if (pdfSet == true) {
		// 인쇄용
		pdfPreset = app.pdfExportPresets.item("MOBIS_인쇄_NEW");
	} else {
		// 검토용
		pdfPreset = app.pdfExportPresets.item("MOBIS_검토_NEW");
		
		// if (pageType == true) { //스프레드 형식
		// 	pdfPreset.exportReaderSpreads = true;
		// } else { //페이지 형식
		// 	pdfPreset.exportReaderSpreads = false;
		// }
	}
	var webMode = false;
	if (myBookContents[0].name.indexOf("000_Cover") != -1) {
		var myFile, coverFile, insideFile;
		// 커버 파일 출력
		// 보안용일 경우 페이지 형식으로 출력, Cover 파일이 있는지 확인할 것
		if (mode == "none" && preset == false) {
			pdfPreset.exportReaderSpreads = false; // 페이지 형식
		} else {
			pdfPreset.exportReaderSpreads = true; // 스프레드 형식
		}
		var coverPDFfName = File.saveDialog('커버 PDF 파일명을 입력하세요', '*.pdf');
		if (coverPDFfName == null) {
			exit();
		}
		coverFile = new File(coverPDFfName);
		app.activeBook.exportFile(ExportFormat.pdfType, coverFile, false, pdfPreset, [myBookContents[0]]);
	} else { webMode = true; }
	
	// 내지를 내보낸다.
	if (colorMode == 0) { // 4도
		pdfPreset.exportReaderSpreads = false; // 페이지 형식
		var insideDocu = [];
		
		if (webMode == true) {
			for (var b=0; b<myBookContents.length; b++) {
				insideDocu.push(myBookContents[b]);
			}
		} else {
			for (var b=1; b<myBookContents.length; b++) {
				insideDocu.push(myBookContents[b]);
			}
		}
		
		var insidePDFfName = File.saveDialog('내지 PDF 파일명을 입력하세요', '*.pdf');
		if (insidePDFfName == null) {
			exit();
		}
		insideFile = new File(insidePDFfName);
		app.activeBook.exportFile(ExportFormat.pdfType, insideFile, false, pdfPreset, insideDocu);
	} else { // 1도
		var prtPreset;
		if (app.activeBook.name.indexOf("MID_AR") != -1) {
			prtPreset = app.printerPresets.item("Mobis_Gray_148x210_Printed_RtL_NEW");
		} else {
			prtPreset = app.printerPresets.item("Mobis_Gray_148x210_Printed_NEW");
		}
		// 커버가 있을 경우 커버 파일을 뺏다가 넣는다.
		var coverfN = myBookContents[0].fullName;
		myBook.bookContents.firstItem().remove();
		myBook.print(false, prtPreset);
		var addfile = File (coverfN);
		myBook.bookContents.add(addfile);
		myBook.bookContents[-1].move(LocationOptions.AT_BEGINNING,myBook.bookContents[0]);
	}

	// } else { //검토용
	// 	var targetPDFfName = File.saveDialog('PDF 파일명을 입력하세요', '*.pdf');
	// 	if (targetPDFfName == null) {
	// 		exit();
	// 	}

	// 	myFile = new File(targetPDFfName);
	// 	app.activeBook.exportFile(ExportFormat.pdfType, myFile, false, pdfPreset);
	// }
	
	// var docs = app.documents;
	// for (var i = docs.length-1; i >= 0; i--) {
	// 	docs[i].close(SaveOptions.YES);
	// }

	alert("완료합니다.");
}

function facingPagesDefault() {
	var doc = app.documents;
	for (var i=0; i<doc.length; i++) {
		if (doc[i].name.indexOf("Cover") != -1) {
			continue
		} else {
			doc[i].documentPreferences.facingPages = true;
		}
	}
}

function visibleAllLayers(docCount) {
	var doc = app.documents;
	
	var slayers;
	// 모든 레이어을 Visible 설정 후 저장
	for (var i=0; i<docCount; i++) {
		slayers = doc[i].layers;
		for (var j=0; j<slayers.length; j++) {
			if (slayers[j].name == "규격") {
				slayers[j].visible = false;
			} else if (slayers[j].name == "#레이어1") {
				slayers[j].visible = true;
			} else {
				if (slayers[j].visible == false) {
					slayers[j].visible = true;
				}
			}
		}
	}
}

function visibleSelectedLayers(docCount) {
	var doc = app.documents;
	var layersName = [];

	// 문서에서 레이어 이름 가져오기
	for (var n=0; n<docCount; n++) {
		getLayersName(doc[n], layersName);
	}
	var layersArray = duplicateRemove(layersName).sort();

	var myDiag = new Window("dialog", "레이어 선택하기");
	myDiag.alignChildren = "left";
	var checkGroup = myDiag.add("panel");
	// var slayers = doc.layers;
	for (var i=0; i<layersArray.length; i++) {
		if (layersArray[i].indexOf("레이어1") != -1 || layersArray[i].indexOf("규격") != -1) {
			continue;
		} else {
			checkGroup.add("checkbox", undefined, layersArray[i]);
		}
	}
	myDiag.add("button", undefined, "OK");

	var result = myDiag.show();

	if (result == 1) {
		var checkedbox = checkGroup.children;
		var sCount = 0;
		var visibleitems = [];
		for (var j=0; j<checkedbox.length; j++) {
			if (checkedbox[j].value == true) {
				sCount ++;
				visibleitems.push(checkedbox[j].text);
			}
		}
		if (sCount == 0) {
			// alert("1개 이상의 레이어를 선택하세요.");
			hideLayersAll();
		} else {
			for (i=0; i<docCount; i++) {
				var myLayer = doc[i].layers;
				for (j=0; j<myLayer.length; j++) {
					var curLayer = myLayer[j];
					for (var n=0; n<sCount; n++) {
						if (visibleitems.toString().indexOf(curLayer.name) != -1) {
							curLayer.visible = true;
						} else if (curLayer.name == "#레이어1") {
							curLayer.visible = true;
						} else if (curLayer.name == "규격") {
								curLayer.visible = false;
						} else {
							curLayer.visible = false;
						}
					}
				}
			}
		}
	} else exit();
}

function hideLayersAll() {
	var doc = app.documents;
	for (var i=0;i<doc.length;i++) {
		slayers = doc[i].layers;
		for (var j=0; j<slayers.length; j++) {
			if (slayers[j].name == "규격") {
				slayers[j].visible = false;
			} else if (slayers[j].name == "#레이어1") {
				slayers[j].visible = true;
			} else {
				if (slayers[j].visible == true) {
					slayers[j].visible = false;
				}
			}
		}
	}
}

function checkBlankPages(doc) {
    var removePage = true;
	var pages = doc.pages.everyItem().getElements();
	var ckPage;
    try {
		if (pages[pages.length - 1].appliedMaster.name == "CH-Memo") {
			ckPage = pages.length - 2;
		} else {
			ckPage = pages.length - 1;
		}
		
		var items = pages[ckPage].pageItems.everyItem().getElements();
		if (pages[ckPage].appliedMaster.name == "CH-Memo") {
			removePage = false;
		} else {
			for(var j=0;j<items.length;j++) {
				if (!(items[j] instanceof TextFrame)) {
					removePage = false;
					break
				}
				if (items[j].contents!="") {
					removePage = false;
					break
				}
			}
		}
		
		if (removePage) {
			alert(ckPage + 1 + " of " + pages.length + " - " + doc.name + " 문서에 빈 페이지가 있습니다.");
			app.activeDocument = doc;
			exit();
		}
	} catch(e) {
		alert(doc.name + "-" + e.line + ":" + e);
		exit();
	}
}

function _checkImages(doc) {
	var Links = doc.links;
	for (var n=0;n<Links.length;n++) {
		if (Links[n].status == 1819109747) {
			app.activeDocument = doc;
			alert(doc.name + " 문서에서 이미지 파일의 유실이 있습니다.\r이미지 유실을 확인한 다음에 다시 실행하세요.");
			exit();
		} else {
			Links[n].update();
		}
	}
}

function _checkCross(doc) {
	var myRef = doc.hyperlinks.everyItem().getElements();
	for (var i=0; i<myRef.length; i++) {
		if (myRef[i].source.constructor == CrossReferenceSource) {
			if (myRef[i].destination == null) {
				app.activeDocument = doc;
				app.activeWindow.activePage = myRef[i].source.sourceText.parentTextFrames[0].parentPage;
				myRef[i].source.sourceText.select();
				alert(doc.name + " 문서에서\r" + myRef[i].source.sourceText + " 상호 참조 오류가 있습니다.");
				exit();
			}
		}
	}
}