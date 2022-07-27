#targetengine "session";
#include 'functions/restix.jsx'
#include 'functions/json2.jsx'

var w = new Window("palette", "AST | 바코드 생성기")
var btn01 = w.add("button", [0, 0, 110, 25], "QR 코드 생성하기");
var btn02 = w.add("button", [0, 0, 110, 25], "코드128 생성하기");

w.show();

btn01.onClick = function() { // QR 코드 생성
	if (app.documents.length == 0) {
		alert("열려있는 문서가 없습니다.");
		exit();
	}
	
	var qrCodeName = prompt("QR 코드의 내용을 입력하세요.", "", "바코드 입력");

	if (qrCodeName == "") {
		alert("입력한 내용이 없습니다.");
		exit();
	}
	if (qrCodeName != null) {
		var myFolder = Folder.selectDialog("QR 코드를 저장할 폴더를 선택하세요.");
	} else {
		exit();
	}
	
	if (myFolder == null) {
		exit();
	}
	
	var myDoc = app.activeDocument;
	// var docPath = myDoc.filePath;
	// QR Code
	var request = {
		url:"https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=" + qrCodeName + "&ecc=H&format=eps"
	}
	
	var outputfile = new File(myFolder + "/qr-" + qrCodeName + ".eps");
	var response = restix.fetchFile(request, outputfile);

	if (response.error) {
		alert("Error, QR 코드를 생성하지 못했습니다.");
		exit();
	}
	alert("생성한 QR 코드를 문서에 넣을 수 있습니다. 현재 선택된 개체가 없다면 문서의 아무 곳에 클릭하세요.");
	myDoc.place(outputfile);
}

btn02.onClick = function() { // 128 코드 생성
	if (app.documents.length == 0) {
		alert("열려있는 문서가 없습니다.");
		exit();
	}
	
	var Code128Name = prompt("128코드의 내용을 입력하세요.", "", "바코드 입력");

	if (Code128Name == "") {
		alert("입력한 내용이 없습니다.");
		exit();
	}
	
	if (Code128Name != null) {
		var myFolder = Folder.selectDialog("128코드를 저장할 폴더를 선택하세요.");
	} else {
		exit();
	}
	
	if (myFolder == null) {
		exit();
	}
	
	var myDoc = app.activeDocument;
	var request = {
		url:"http://www.barcode-generator.org/zint/api.php?bc_number=20&bc_data=" + Code128Name + "&bc_download=1&bc_format=2&bc_size=l"
	}
	
	var outputfile = new File(myFolder + "/" + Code128Name + ".eps");
	var response = restix.fetchFile(request, outputfile);

	if (response.error) {
		alert("Error, 128코드를 생성하지 못했습니다.");
		exit();
	}
	alert("생성한 128코드를 문서에 넣을 수 있습니다. 현재 선택된 개체가 없다면 문서의 아무 곳에 클릭하세요.");
	myDoc.place(outputfile);
}