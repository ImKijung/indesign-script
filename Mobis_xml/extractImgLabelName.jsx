// 북파일이 열려있는지 확인한다.
if (app.books.length == 0) {
	alert("북 파일이 열려있지 않습니다.");
	exit();
}
var myBook = app.activeBook;
var bookPath = myBook.filePath;

// Report 파일을 생성한다.
var Report = new File(bookPath +"/" + myBook.name.replace(".indb","") + "_label.txt");
Report.open("w");
Report.encoding = "utf-8";

// 각 파일을 열어서 이미지 파일명을 Report 파일에 기록한다.
var myBookContents = myBook.bookContents.everyItem().getElements();
for (var b=0; b<myBookContents.length; b++) {
	var myPath = myBookContents[b].fullName;
	var myDoc = File(myPath);
	app.open(myDoc);
	eXtractLabelName(myDoc, Report);
}
Report.close();


function eXtractLabelName(doc, Report) {
	doc = app.activeDocument;
	var myLinks = doc.links;
	var myLinkName;
	for (var i=0; i<myLinks.length; i++) {
		myLinkName = myLinks[i].name;
		if (myLinkName.indexOf("D-") != -1) {
			Report.writeln(doc.name + "\t" + myLinkName);
		}
	}
	// 파일을 닫는다.
	doc.close();
}