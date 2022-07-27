#targetengine "session";
#include 'functions/restix.jsx'
#include 'functions/json2.jsx'

var w = new Window("palette", "AST > MA | 바코드 생성기 ver.1.0.2")
var btn01 = w.add("button", [0, 0, 225, 25], "QR 코드 생성하기");
var btn02 = w.add("button", [0, 0, 225, 25], "코드128 생성하기");
var btn03 = w.add("button", [0, 0, 225, 25], "Data Matrix 생성하기");

w.show();

btn01.onClick = function() { // QR 코드 생성
	w.close();
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
		url:"https://api.qrserver.com/v1/create-qr-code/?size=1000x1000&data=" + qrCodeName + "&ecc=L&format=png"
	}
	
	var outputfile = new File(myFolder + "/" + "E-usermanual_XXXX_eng.png");
	var response = restix.fetchFile(request, outputfile);

	if (response.error) {
		alert("Error, QR 코드를 생성하지 못했습니다.");
		exit();
	}
	savePNG(outputfile);

	function savePNG(outputfile) {
		var myFile = new File(outputfile);
		var bt = new BridgeTalk();
		bt.target = "photoshop";
		bt.body = "var myDoc = app.open(new File('" + myFile + "'));";
		bt.body += "myDoc.changeMode(ChangeMode.GRAYSCALE);";
		// bt.body += "myDoc.resizeImage('23mm', '23mm', 300);";
		// bt.body += "pngOptions = new PNGSaveOptions();";
		// bt.body += "pngOptions.compression = 0;";
		// bt.body += "pngOptions.interlaced = false;";
		// bt.body += "savePath = File(myDoc.path + '/' + myDoc.name.replace('-down', ''));"
		// bt.body += "myDoc.saveAs(savePath, pngOptions, false, Extension.LOWERCASE);";
		bt.body += "myDoc.save();";
		bt.body += "myDoc.close();";
		// bt.body += psScript().toString() + "psScript();";
		// $.writeln(bt.body);
		bt.onResult = function(resObj) {
			var myResult = resObj.body;
		}
		bt.onError = function( inBT ) { alert(inBT.body); };
		bt.send(8);
	}

	alert("QR 코드 생성을 완료합니다.");
	
	w.show();
}

btn02.onClick = function() { // 128 코드 생성
	w.close();
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
		url:"http://www.barcode-generator.org/zint/api.php?bc_number=20&bc_data=" + Code128Name + "&bc_download=1&bc_format=2&bc_size=m"
	}
	
	var outputfile = new File(myFolder + "/" + Code128Name + ".eps");
	var response = restix.fetchFile(request, outputfile);

	if (response.error) {
		alert("Error, 128코드를 생성하지 못했습니다.");
		exit();
	}
	alert("생성한 128코드를 문서에 넣을 수 있습니다. 현재 선택된 개체가 없다면 문서의 아무 곳에 클릭하세요.");
	myDoc.place(outputfile);
	w.show();
}

btn03.onClick = function() {
	w.close();
	app.scriptPreferences.measurementUnit = MeasurementUnits.MILLIMETERS;

	if (app.documents.length == 0) {	
		alert("열려있는 문서가 없습니다.");
		exit();
	}
	
	var dmxName = prompt("Data Matrix 코드의 내용을 입력하세요.", "", "바코드 입력");

	if (dmxName == "") {
		alert("입력한 내용이 없습니다.");
		exit();
	}
	
	if (dmxName != null) {
		var myFolder = Folder.selectDialog("Data Matrix 코드를 저장할 폴더를 선택하세요.");
	} else {
		exit();
	}
	
	if (myFolder == null) {
		exit();
	}
	dmxName = dmxName.replace(" ", "+");
	var myDoc = app.activeDocument;
	var request = {
		url:"http://generator.onbarcode.com/datamatrix.aspx?DATA=" + dmxName + "&PROCESS-TILDE=false&DATA-MODE=0&FORMAT-MODE=0&FNC1=0&UOM=0&X=3&LEFT-MARGIN=0&RIGHT-MARGIN=0&TOP-MARGIN=0&BOTTOM-MARGIN=0&RESOLUTION=96&ROTATE=0&BARCODE-WIDTH=0&BARCODE-HEIGHT=0&IMAGE-FORMAT=png"
	}
	var outputfile = new File(myFolder + "/" + dmxName.replace("+", " ") + ".png");
	var response = restix.fetchFile(request, outputfile);

	if (response.error) {
		alert("Error, Data Matrix 코드를 생성하지 못했습니다.");
		exit();
	}
	savePNG(outputfile);

	// var dmxFrame = myDoc.rectangles.add({
	// 		geometricBounds:[0, 0, 7.5, 7.5],
	// 		contentType:ContentType.GRAPHIC_TYPE
	// 	});
	// dmxFrame.fillColor = myDoc.swatches[2];
	// dmxFrame.place(outputfile);
	// dmxFrame.fit(FitOptions.CONTENT_TO_FRAME);
	// dmxFrame.geometricBounds = [0, 0, 9.5, 9.5];
	// dmxFrame.fit(FitOptions.CENTER_CONTENT);
	w.show();

	function savePNG(outputfile) {
		var myFile = new File(outputfile);
		var bt = new BridgeTalk();
		bt.target = "photoshop";
		bt.body = "var myDoc = app.open(new File('" + myFile + "'));";
		bt.body += "myDoc.changeMode(ChangeMode.GRAYSCALE);";
		bt.body += "myDoc.save();";
		bt.body += "myDoc.close();";
		// bt.body += psScript().toString() + "psScript();";
		// $.writeln(bt.body);
		bt.onResult = function(resObj) {
			var myResult = resObj.body;
		}
		bt.onError = function( inBT ) { alert(inBT.body); };
		bt.send(8);
	}
}