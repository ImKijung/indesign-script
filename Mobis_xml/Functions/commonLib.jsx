function duplicateRemove(layersName) { // 레이어 중복 제거
	var m = {};
	var newarr = [];
	for (var x=0; x<layersName.length; x++) {
		var v = layersName[x];
		if (!m[v]) {
			newarr.push(v);
			m[v] = true;
		}
	}
	return newarr
}

function getLayersName(doc, layersName) { // 레이어 이름 가져오기
	var ilayers = doc.layers;
	for (var a=0; a<ilayers.length; a++) {
		layersName.push(ilayers[a].name);
	}
	return layersName;
}

function checkBookFileOrder() {
	var fileList = [];
	var myBook = app.activeBook;
	var myBookContents = myBook.bookContents.everyItem().getElements();
	for (var i=0; i<myBookContents.length; i++) {
		var myDocName = myBookContents[i].name;
		fileList.push(myDocName);
	}
	var chpNum
	var orderList = [];
	for (i=0; i<fileList.length; i++) {
		chpNum = fileList[i].replace(/(\_|\-)[^>]+\.indd/ig, "")
		$.writeln(chpNum);
		orderList.push(chpNum*1);
	}
	for (i=0; i<orderList.length-1; i++) {
		if (orderList[i] <= orderList[i+1]) {
			continue;
		} else {
			alert ("파일명 " + orderList[i] + " 파일의 순서가 잘못됐습니다. 북 파일을 확인하세요.");
			exit();
		}
	}
	alert("파일명 순서가 올바릅니다.");
}