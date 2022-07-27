#targetengine "session"

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
// alert("생성한 Data Matrix 코드를 문서에 넣을 수 있습니다. 현재 선택된 개체가 없다면 문서의 아무 곳에 클릭하세요.");

savePNG(outputfile);

var dmxFrame = myDoc.rectangles.add({
		geometricBounds:[0, 0, 7.5, 7.5],
		contentType:ContentType.GRAPHIC_TYPE
	});
dmxFrame.fillColor = myDoc.swatches[2];
dmxFrame.place(outputfile);
dmxFrame.fit(FitOptions.CONTENT_TO_FRAME);
dmxFrame.geometricBounds = [0, 0, 9.5, 9.5];
dmxFrame.fit(FitOptions.CENTER_CONTENT);

function savePNG(outputfile) {
	var myFile = new File(outputfile);
	var bt = new BridgeTalk();
	bt.target = "photoshop";
	bt.body = "var myDoc = app.open(new File('" + myFile + "'));";
	bt.body += "myDoc.changeMode(ChangeMode.GRAYSCALE);";
	bt.body += "myDoc.save();";
	bt.body += "myDoc.close();";
	// bt.body += psScript().toString() + "psScript();";
	$.writeln(bt.body);
	bt.onResult = function(resObj) {
		var myResult = resObj.body;
	}
	bt.onError = function( inBT ) { alert(inBT.body); };
	bt.send(8);
}

