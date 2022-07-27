if (app.books.length == 0) {
	alert("북 파일이 열려 있지 않습니다.");
	exit();
}

var myBook = app.activeBook;
var lastPages = null;
var myBookContents = myBook.bookContents.everyItem().getElements();

var myDlg = new Window ("dialog", "페이지 번호 입력");
var PageNumber = myDlg.add('edittext {characters: 10, justyfy: "left"}', undefined, "");
PageNumber.active = true;	
myDlg.add("button", undefined, "OK");
var result = myDlg.show();
var myPath, myDoc;

if (result == 1) {
	var booksPage = PageNumber.text.split("-");
	if (booksPage[0] == "I") {
		for (var i=0; i<myBookContents.length; i++) {
			if (myBookContents[i].name.indexOf("INDEX") != -1) {
				myPath = myBookContents[i].fullName;
				myDoc = File(myPath);
			}
		}
	} else {
		myPath = myBookContents[booksPage[0]-1].fullName;
		myDoc = File(myPath);
	}

	app.open(myDoc);
	var doc = app.activeDocument;
	var lastPageNum = doc.pages.lastItem().name;
	// $.writeln(lastPageNum + ":" + PageNumber.text);

	try {
		if (parseInt(booksPage[1]) > parseInt(lastPageNum.split("-")[1])) {
			alert(PageNumber.text + " 입력한 페이지 번호가 현재 열려있는 문서의 페이지 범위에 포함하지 않습니다.");
		} else {
			var selectedPage = doc.pages.itemByName(PageNumber.text);
			app.activeWindow = doc.layoutWindows[0];
			app.activeWindow.activePage = selectedPage;
		}
	} catch (e) {
		alert("입력한 페이지를 찾을 수 없습니다.");
	}
	
} else exit();