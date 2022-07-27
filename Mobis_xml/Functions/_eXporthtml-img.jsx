// eXportHTMLimg();

function eXportHTMLimg() {
    // var doc = app.activeDocument;
	var myBook = app.activeBook;
	var curhtmlfp = Folder(myBook.filePath + "/html_img");
	if (curhtmlfp.exists) {
		removeFiles(curhtmlfp);
		curhtmlfp.remove();
	}
	var myBookContents = myBook.bookContents.everyItem().getElements();
	var myResult, doc;
	// 경고창 비활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	// 프로그레스 바
	var myProg = new Window ('palette', 'HTML 이미지 생성 시작');
		myProg.pbar = myProg.add ('progressbar', undefined, 0, myBookContents.length);
	myProg.pbar.preferredSize.width = 300;
	myProg.show();

	for (var b=0; b<myBookContents.length; b++) {
		var myPath = myBookContents[b].fullName;
		var myDoc = File(myPath);
		myProg.pbar.value = b + 1;
		app.open(myDoc);
		doc = app.activeDocument;
		myProg.text = doc.name + ' 이미지 상태 검사 중...';
		myResult = _checkimgstatus(doc);
		if (myResult == -1) {
			return false;
		}
		myProg.text = doc.name + ' HTML 이미지 생성 중...';
		_makeHTMLimg(doc);
		doc.close(SaveOptions.YES);
	}
	
	myProg.close();
	// 경고창 활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
	alert("완료합니다.");
}

function _checkimgstatus(doc) {
	var link, newName, curLinkFile, pngName, pngFile
	for (var d = doc.links.length-1; d >= 0; d--) {
		link = doc.links[d];
		if (link.status == LinkStatus.LINK_OUT_OF_DATE || link.status == LinkStatus.LINK_MISSING) {  
			return -1
		}
		if (link.name.indexOf(".PNG") != -1) {
			newName = link.name.replace(".PNG", ".png");
			curLinkFile = new File(link.filePath);
			curLinkFile.rename(newName);
			link.relink(curLinkFile);
		}
		if (link.name.indexOf(".jpg") != -1 || link.name.indexOf(".JPG") != -1) {
			pngName = link.filePath.replace(/\.(jpg|JPG)/gi, ".png");
			pngFile = new File(pngName);
			link.parent.exportFile(ExportFormat.PNG_FORMAT, pngName);
			link.relink(pngFile);
		}
   }
}

function _makeHTMLimg(doc) {
	var docPath = doc.filePath;
    var imgCount = doc.allGraphics.length;

    var htmlPath = new Folder(docPath + "/html_img");
    if (!htmlPath.exists) {
        htmlPath.create();
    } else {
		htmlPath.create();
	}

    app.pngExportPreferences.exportResolution = 300;
    app.pngExportPreferences.pngQuality = PNGQualityEnum.HIGH;
    app.pngExportPreferences.transparentBackground = true;

    var linkName, pngName

    for (var i=0;i<imgCount;i++) {
        linkName = doc.allGraphics[i].itemLink.name;
        pngName = new File(htmlPath + "/" + linkName.replace(".ai", "").replace(".png", "").replace(".PNG", "") + ".png");
        doc.allGraphics[i].exportFile(ExportFormat.PNG_FORMAT, pngName);
    }
}

function removeFiles(folder) {
    var files = folder.getFiles();
    var n
    for (n=files.length - 1; n >= 0; n--){
        files[n].remove();
	};
};