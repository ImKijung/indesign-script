app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

if (app.books.length == 0) {
	alert("인디자인 북 파일을 열어주세요.");
	exit();
} else if (app.books.length > 0) {
	var myBook = app.activeBook;
	var myBookContents = myBook.bookContents.everyItem().getElements();
	for(var i=0; i<myBookContents.length; i++) {
		var myPath = myBookContents[i].fullName;
		var myFile = File(myPath);
		var doc = app.open(myFile);
	}
}
var docs = app.documents;
var myWindow = new Window ('palette', '텍스트 넘침 확인 중 ...');
	myWindow.pbar = myWindow.add ('progressbar', undefined, 0, docs.length);
	myWindow.pbar.preferredSize.width = 300;
	myWindow.show();
for (j=0; j<docs.length; j++) {
	myWindow.pbar.value = j+1;
	var openDocs = app.documents.everyItem().getElements();
	app.activeDocument = openDocs[openDocs.length-1];
	checkOverFlow();
	app.activeDocument.save();
	$.sleep(20); // Do something useful here
}
var newLinkpath = Folder.selectDialog( "이미지 경로를 선택하세요." );
if (newLinkpath == null) {
	exit();
}
var iiWindow = new Window ('palette', '이미지 경로 재설정 중 ...');
	iiWindow.pbar = iiWindow.add ('progressbar', undefined, 0, docs.length);
	iiWindow.pbar.preferredSize.width = 300;
	iiWindow.show();
for (j=0; j<docs.length; j++) {
	iiWindow.pbar.value = j+1;
	var openDocs = app.documents.everyItem().getElements();
	app.activeDocument = openDocs[openDocs.length-1];
	replaceAllimagees(newLinkpath);
	app.activeDocument.save();
	$.sleep(20); // Do something useful here
}

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
alert("완료");

function checkOverFlow() {
	// Overflow - script for finding overflowed text in InDesign
	// Pawel Swiecicki 22.08.2006
	var of = new Array();
	var nodocs = "열려있는 문서가 없습니다.";
	var noframes = "현재 문서에는 텍스트 프레임이 없습니다.";
	var onpages = "";
	var founded = "넘치는 텍스트 프레임을 찾았습니다.: ";
	var oop = "페이지: ";
	var notfound = "넘치는 텍스트 프레임이 없습니다.";

	if (app.documents.length < 1){  alert(nodocs);}
	else {
		if (app.activeDocument.textFrames.length < 1) {alert(noframes);}
	else {
		for(var i = 0; i < app.activeDocument.pages.count(); i++) {
			for(var b = 0; b < app.activeDocument.pages[i].textFrames.count(); b++) { 
				if(app.activeDocument.pages[i].textFrames[b].overflows == true) {
					of[of.length] = app.activeDocument.pages[i].name;
				}
			}
		}
		if(of.length > 0) {
			for(i = 0; i < of.length; i++) {
				if(i > 0) {
					onpages += ", ";
				}
				onpages += of[i];
			}
			alert(founded + of.length + "\n\n" + oop + onpages + "\r다음 페이지를 확인하세요.");
			app.activeDocument.pages[-1].select();
			exit();
		} else {
			// alert(notfound);
		}
		}
	}
}

function replaceAllimagees(newLinkpath) {
	var doc = app.activeDocument;
	app.pdfPlacePreferences.pdfCrop = PDFCrop.CROP_BLEED;

	var myGraphics = doc.allGraphics;
	var imageCount = GetImageCountWithoutMaster(doc);

	for (var i = 0; i < myGraphics.length; i++) {
		var myGraphic = myGraphics[i];
		var myPath = myGraphic.itemLink.filePath;
		var myfName = myGraphic.itemLink.name;
		var imageFile = File(newLinkpath + "/" + myfName)
		// $.writeln(imageFile);
		// $.writeln(imageFile.exists);
		if (!imageFile.exists) {
			alert("다음 경로의 파일이 없습니다. 파일을 확인한 다음 스크립트를 다시 실행하세요.\r" + imageFile);
			exit();
		} else
		try {
			if (myGraphic.constructor.name == "PDF") {
				myGraphic.place(imageFile, false);
				// myGraphics[i].horizontalScale = 100;
				// myGraphics[i].verticalScale = 100;
				// myGraphics[i].parent.fit(FitOptions.FRAME_TO_CONTENT);
			} else
			myGraphic.place(imageFile, false);
		} catch(e) {
			alert(e + ":" + e.line + " - " + myfName);
			exit();
		}
	}

	var newGraphics = doc.allGraphics;

	for (var j = 0; j < imageCount; j++) {
		var newGraphic = newGraphics[j];
		newGraphic.horizontalScale = 100;
		newGraphic.verticalScale = 100;
		newGraphic.parent.fit(FitOptions.FRAME_TO_CONTENT);
	}
}

function GetImageCountWithoutMaster(myDoc) {
	try {
		var myMaster = myDoc.masterSpreads;
		var myMasterNumber = myMaster.count();
		var masterImages = 0;
		for (var i = 0 ; i < myMasterNumber ; i++) {
			imageNumber = myMaster[i].allGraphics;
			masterImages += imageNumber.length;
		}

		return myDoc.allGraphics.length - masterImages;
	} catch (ex) {
		alert(ex);
	}
}