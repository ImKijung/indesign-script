var closeBook = "열려있는 북 파일을 닫고 실행하세요.";
var noBook = "열려있는 북 파일이 없습니다.";
var closeDocu = "열려있는 문서를 닫고 실행하세요.";
var noDocu = "열려있는 문서가 없습니다.";

function preCheck(docu) {
	if (docu.constructor.name == "Book") {
		var Book = app.activeBook;
		CheckBook(Book);
	} else if (docu.constructor.name == "Document") {
		var doc = app.activeDocument;
		
	}
}

function openDocuments() {
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	var myBook = app.activeBook;
	var bookContent;
	for (var i=0; i<myBook.bookContents.length; i++) {
		bookContent = myBook.bookContents[i];
		curDocFile = new File(myBook.filePath + "/" + bookContent.name);
		app.open(curDocFile);
	}
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
}

function closeDocs() {
	docs = app.documents;
	for (var i=docs.length-1;i>=0;i--) {
		docs[i].close(SaveOptions.NO);
	}
}

function CheckBook(Book) {
	var missingDocs = [];
	var bookContent;
	for (var i=0; i<Book.bookContents.length; i++) {
		bookContent = Book.bookContents[i];
		if (bookContent.status == BookContentStatus.MISSING_DOCUMENT) {
			missingDocs.push(bookContent.name);
		}
	}

	if (missingDocs != 0) ErrorExit(missingDocs.length + " 개의 문서가 북 파일과 연결되어 있지 않습니다." + ((missingDocs.length === 1) ? ": " : "s: ") + missingDocs.join(", "), true);
}

function eXportIDML(oPath, doc) {
	var oIDMLf = new File(oPath + "/" + doc.name.replace('.indd', '.idml'));
	doc.exportFile(ExportFormat.INDESIGN_MARKUP, oIDMLf);
}

function getData() {
	var pdfPresets = app.pdfExportPresets.everyItem().name;
	var dlg = app.dialogs.add( { name : 'Choose Style' } );
	with(dlg) {
		with( dialogColumns.add() ) {
			with( borderPanels.add( ) ) {
				staticTexts.add( { staticLabel : 'PDF Presets:', minWidth : 93 } );
				dropDown = dropdowns.add( { stringList : pdfPresets, selectedIndex : 0 });
			}
		}
	}
	if( dlg.show() == false ){
		dlg.destroy();
		exit(0);
	}
	return dropDown.selectedIndex;
}

function ErrorExit(error, icon) {
	alert(error, icon);
	exit();
}

function inputVersion() {
	var ver = prompt ("", "D", "판차 정보를 입력하세요.");
	if (ver == null) {
		exit();
	}
	return ver
}

function GetDate() {
	var date = new Date();
	var year, month, day;
	if ((date.getYear() - 100) < 10) {
		year = "0" + new String((date.getYear() - 100));
	} else {
		year = new String((date.getYear() - 100));
	}
	if (date.getMonth() + 1 < 10) {
		month =  "0" + new String(date.getMonth() + 1);
	} else {
		month =  new String(date.getMonth() + 1);
	}
	if (date.getDate() < 10) {
		day =  "0" + new String(date.getDate());
	} else {
		day =  new String(date.getDate());
	}
	var dateString = year + month + day;
	return dateString;
}

function GetAppVersion() {
	var appVersion,
	appVersionNum = Number(String(app.version).split(".")[0]);
	
	switch (appVersionNum) {
		case 17:
			appVersion = "2022";
			break;
		case 16:
			appVersion = "2021";
			break;
		case 15:
			appVersion = "2020";
			break;	
		case 14:
			appVersion = "CC 2019";
			break;
		case 13:
			appVersion = "CC 2018";
			break;
		case 12:
			appVersion = "CC 2017";
			break;
		case 11:
			appVersion = "CC 2015";
			break;
		case 10:
			appVersion = "CC 2014";
			break;			
		case 9:
			appVersion = "CC";
			break;		
		case 8:
			appVersion = "CS 6";
			break;
		case 7:
			if (app.version.match(/^7\.5/) != null) {
				appVersion = "CS 5.5";
			}
			else {
				appVersion = "CS 5";
			}
			break;
		case 6:
			appVersion = "CS 4";
			break;
		case 5:
			appVersion = "CS 3";
			break;
		case 4:
			appVersion = "CS 2";
			break;
		case 3:
			appVersion = "CS";
			break;			
		default:
		return null;
	}

	return appVersion;
}