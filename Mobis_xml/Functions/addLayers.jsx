#include "commonLib.jsx";

function AddLayer(docuType) {
	var docCount;
	if (docuType == "activeDocument") {
		docCount = 1;
		AddLayerMain(docCount);
	} else if (docuType == "documents") {
		docCount = app.documents.length;
		AddLayerMain(docCount);
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
		AddLayerMain(docCount);
	}
}

function AddLayerMain(docCount) {
	var doc = app.documents;
	var layersName = [];

	for (var n=0; n<docCount; n++) {
		getLayersName(doc[n], layersName);
	}

	var layersArray = duplicateRemove(layersName).sort();
	var newLayerName = prompt("현재 레이어 : " + layersArray + "\r추가할 레이어 이름을 입력하세요.", "", "레이어 추가하기");

	if (newLayerName == "") {
		alert("레이어 이름을 입력하세요.")
		exit();
	}
	// 레이어 색상 설정
	var lColor;
	if (newLayerName.indexOf("1차") != -1) {
		lColor = UIColors.RED;
	} else if (newLayerName.indexOf("2") != -1) {
		lColor = UIColors.GREEN;
	} else if (newLayerName.indexOf("3") != -1) {
		lColor = UIColors.BLUE;
	} else if (newLayerName.indexOf("4") != -1) {
		lColor = UIColors.YELLOW;
	} else if (newLayerName.indexOf("5") != -1) {
		lColor = UIColors.MAGENTA;
	} else if (newLayerName.indexOf("6") != -1) {
		lColor = UIColors.CYAN;
	} else if (newLayerName.indexOf("7") != -1) {
		lColor = UIColors.GRAY;
	} else if (newLayerName.indexOf("8") != -1) {
		lColor = UIColors.BLACK;
	} else if (newLayerName.indexOf("9") != -1) {
		lColor = UIColors.ORANGE;
	} else if (newLayerName.indexOf("0") != -1) {
		lColor = UIColors.DARK_GREEN;
	} else if (newLayerName.indexOf("음영") != -1) {
		lColor = UIColors.SULPHUR;
	} else if (newLayerName.indexOf("도안번호") != -1) {
		lColor = UIColors.LIPSTICK;
	}

	var errCount = 0;
	var errfilelist = "";
	if (docCount > 1) {
		var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, docCount);
		myProg.pbar.preferredSize.width = 300;
		myProg.show();
	}
	for (n=0; n<docCount; n++) {
		if (docCount > 1) {
			// 추가 코드
			myProg.text = doc[n].name + ' 레이어 추가...';
			myProg.pbar.value = n + 1;
		}
		var slayers = doc[n].layers;
		var addLayer = slayers.itemByName(newLayerName);
		if (addLayer == null) {
			slayers.add({
				name: newLayerName,
				layerColor: lColor
			}).move(LocationOptions.AT_BEGINNING);
		} else {
			errCount ++;
			errfilelist = errfilelist + doc[n].name + " 문서에는 이미 " + newLayerName + " 레이어가 있습니다.\r";
		}
		doc[n].save();
	}
	if (docCount > 1) {
		// 종료
		myProg.close();
	}
	
	// 같은 레이어 명이 있을 경우 알림 표시
	if (errCount > 0) {
		alert(errfilelist);
	} else {
		alert("완료합니다.");
	}
}