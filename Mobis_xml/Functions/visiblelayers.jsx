#include "hidelayers.jsx";

function VisibleLayers(docuType) {
	var docCount;
	if (docuType == "activeDocument") {
		docCount = 1;
		VisibleLayersMain(docCount);
	} else if (docuType == "documents") {
		docCount = app.documents.length;
		VisibleLayersMain(docCount);
	} else if (docuType == "activeBook") {
		// 북파일
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		for (var b=0; b<myBookContents.length; b++) {
			var myPath = myBookContents[b].fullName;
			var myDoc = File(myPath);
			app.open(myDoc);
		}
		docCount = app.documents.length;
		VisibleLayersMain(docCount);
	}
}

function VisibleLayersMain(docCount) {
	var doc = app.documents;
	// var ilayers = doc.layers;
	var layersName = [];

	// 문서에서 레이어 이름 가져오기
	for (var n=0; n<docCount; n++) {
		getLayersName(doc[n], layersName);
	}
	var layersArray = duplicateRemove(layersName).sort();
	
	// 레이어 이름 정렬하기 
	// $.writeln(layersArray);

	var myDiag = new Window("dialog", "레이어 보이기");
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
			alert("1개 이상의 레이어를 선택하세요.");
		} else {
			if (docCount == 1) { // 현재 활성화된 문서
				for (var k=0; k<sCount; k++) {
					if (app.documents[0].layers.item(visibleitems[k]).isValid) {
						app.documents[0].layers.item(visibleitems[k]).visible = true;
						app.documents[0].save();
					}
				}
			} else if (docCount > 1) { // 열려있는 모든 문서
				var myProg = new Window ('palette', '실행 중...');
					myProg.pbar = myProg.add ('progressbar', undefined, 0, docCount);
				myProg.pbar.preferredSize.width = 300;
				myProg.show();
				for (n=0; n<docCount; n++) {
					myProg.text = app.documents[n].name + ' 레이어 보이기...';
					myProg.pbar.value = n + 1;
					for (k=0; k<sCount; k++) {
						if (app.documents[n].layers.item(visibleitems[k]).isValid) {
							app.documents[n].layers.item(visibleitems[k]).visible = true;
							app.documents[n].save();
						}
					}
				}
				myProg.close();
			}
		}
	} else exit();
}