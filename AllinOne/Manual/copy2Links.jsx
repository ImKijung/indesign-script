#targetengine "session";

// copy2Links();
function copy2Links() {
	var mainW = new Window ("palette", "이미지를 모아 모아");
		mainW.alignChildren = "left";
		var btn_01 = mainW.add ("button", [0,0,250,30], "[1]image - SERIAL2 폴더로");
		var btn_02 = mainW.add ("button", [0,0,250,30], "[1]Links - 인디자인의 Links 폴더로");
		var btn_03 = mainW.add ("button", [0,0,250,30], "[b]image - Book문서의 SERIAL2 폴더로");
		var btn_04 = mainW.add ("button", [0,0,250,30], "[b]Links - Book문서의 Links 폴더로");
		var btn_06 = mainW.add ("button", [0,0,250,30], "[f]Folder - 폴더 선택하기");

	btn_01.onClick = function () {
		mainW.show ();
		var mDoc = app.activeDocument;
		var mLinks = mDoc.links.everyItem().getElements();
		var destFolder = mDoc.filePath.parent + "/image/";
		var len = mLinks.length;
		if ( !Folder(destFolder).create() ) exit();

		while (len-->0) {
			var currLinkFName = mLinks[len].filePath;
			var currFile = File(destFolder + File(currLinkFName).name);
			if (!currFile.exists && File(currLinkFName).exists) {
				mLinks[len].copyLink(currFile);
				} else
			mLinks[len].relink(currFile)
		}
		alert ("이미지가 " + destFolder + " 폴더로 정리되었습니다.");
	}

	btn_02.onClick = function () {
		mainW.show ();
		var mDoc = app.activeDocument;
		var mLinks = mDoc.links.everyItem().getElements();
		var destFolder = mDoc.filePath + "/Links/";
		var len = mLinks.length;

		if ( !Folder(destFolder).create() ) exit();

		while (len-->0) {
			var currLinkFName = mLinks[len].filePath;
			var currFile = File(destFolder + File(currLinkFName).name);
			if (!currFile.exists && File(currLinkFName).exists) {
				mLinks[len].copyLink(currFile);
				} else
			mLinks[len].relink(currFile)
		}
		alert ("이미지가 " + destFolder + " 폴더로 정리되었습니다.");
	}

	btn_03.onClick = function () {
		mainW.show ();
		var myBook = app.activeBook
		var mDocs = myBook.bookContents.everyItem().getElements();
		for (var i=0; i < mDocs.length; i++) {
			var myPath = mDocs[i].fullName;
			var mDoc = File(myPath);
			var myDoc = app.open(mDoc);
			var mLinks = myDoc.links.everyItem().getElements();
			var destFolder = myDoc.filePath.parent + "/image/";
			var len = mLinks.length;
		
		
			if ( !Folder(destFolder).create() ) exit();

			while (len-->0) {
				var currLinkFName = mLinks[len].filePath;
				var currFile = File(destFolder + File(currLinkFName).name);
				if (!currFile.exists && File(currLinkFName).exists) {
					mLinks[len].copyLink(currFile);
					} else
				mLinks[len].relink(currFile)
			}
			myDoc.close(SaveOptions.YES);
		}
		alert ("이미지가 " + destFolder + " 폴더로 정리되었습니다.");
	}

	btn_04.onClick = function () {
		mainW.show ();
		var myBook = app.activeBook
		var mDocs = myBook.bookContents.everyItem().getElements();
		for (var i=0; i < mDocs.length; i++) {
			var myPath = mDocs[i].fullName;
			var mDoc = File(myPath);
			var myDoc = app.open(mDoc);
			var mLinks = myDoc.links.everyItem().getElements();
			var destFolder = myDoc.filePath + "/Links/";
			var len = mLinks.length;
		
		
			if ( !Folder(destFolder).create() ) exit();

			while (len-->0) {
				var currLinkFName = mLinks[len].filePath;
				var currFile = File(destFolder + File(currLinkFName).name);
				if (!currFile.exists && File(currLinkFName).exists) {
					mLinks[len].copyLink(currFile);
					} else
				mLinks[len].relink(currFile)
			}
			myDoc.close(SaveOptions.YES);
		}
		alert ("이미지가 " + destFolder + " 폴더로 정리되었습니다.");
	}

	btn_06.onClick = function () {
		mainW.show ();
		var mDoc = app.activeDocument;
		var mLinks = mDoc.links.everyItem().getElements();
		//var destFolder = mDoc.filePath.parent + "/image/";
		var myFolder = Folder.selectDialog ('이미지를 새롭게 모을 폴더를 선택하세요. Links 폴더를 만들어 저장합니다.');
		var destFolder = myFolder + "/Links/";
		var len = mLinks.length;
		if ( !Folder(destFolder).create() ) exit();

		while (len-->0) {
			var currLinkFName = mLinks[len].filePath;
			var currFile = File(destFolder + File(currLinkFName).name);
			if (!currFile.exists && File(currLinkFName).exists) {
				mLinks[len].copyLink(currFile);
				} else
			mLinks[len].relink(currFile);
		}
		alert ("이미지가 " + destFolder + " 폴더로 정리되었습니다.");
	}

	mainW.show ();
}