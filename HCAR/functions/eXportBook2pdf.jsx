function _eXportBook2pdf() {
	var myFolder = Folder.selectDialog("인디자인 북파일의 상위 폴더를 선택하세요.");

	if (myFolder == null) {
		exit();
	}

	var myFiles = [];
	GetSubBooks(myFolder);

	if (myFiles.length == 0) {
		alert("선택한 폴더 하위에 북 파일이 없습니다.");
		exit();
	}

	var outFolder = Folder.selectDialog("PDF를 저장할 폴더를 선택하세요.");

	if (outFolder == null) {
		alert("PDF를 저장할 위치를 선택하지 않아 프로세스를 종료합니다.");
		exit();
	}

	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	// var prtPreset = app.printerPresets.item("CAR_148x210_WEB");
	// AST_CAR_Manual
	var pdfPreset = app.pdfExportPresets.item("AST_CAR_Manual");
	var myBook, eXpdf, pdfName, myBookContents, myPath, myDoc, docs, outPDF;
	var dt = new Date();
	var month = dt.getMonth() + 1;
	var day = dt.getDate();
	if (month < 10) {
		month = "0" + month.toString();
	}
	if (day < 10) {
		day = "0" + day.toString();
	}
	var today = dt.getFullYear().toString().substr(2) + month + day;

	for (var i=0;i<myFiles.length;i++) {
		myBook = app.open(myFiles[i]);
		
		pdfName = myBook.name.replace('.indb', '_' + today + '.pdf');
		// eXpdf = new File(myBook.filePath + "/" + pdfName);
		// myBook.printPreferences.printFile = eXpdf;

		myBookContents = myBook.bookContents.everyItem().getElements();
		
		for (var j=0;j<myBookContents.length;j++) {
			myPath = myBookContents[j].fullName;
			myDoc = File(myPath);
			app.open(myDoc);
		}
		// myBook.print(false, prtPreset);
		outPDF = new File(outFolder + "/" + pdfName)
		myBook.exportFile(ExportFormat.pdfType, outPDF, false, pdfPreset);
		
		docs = app.documents;
		for (var k = docs.length-1; k >= 0; k--) {
			docs[k].close(SaveOptions.NO);
		}
		myBook.close(SaveOptions.NO);
	}
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

	function GetSubBooks(theFolder) {
		var myFileList = theFolder.getFiles();
		for (var i = 0; i < myFileList.length; i++) {
			var myFile = myFileList[i];
			if (myFile instanceof Folder) {
				GetSubBooks(myFile);
			}
			else if (myFile instanceof File && myFile.name.match(/\.indb$/i)) {
				myFiles.push(myFile);
			}
		}
	}
}