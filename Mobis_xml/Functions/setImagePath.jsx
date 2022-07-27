function rePlaceImagePath(docuType) {
	var myFolder = Folder.selectDialog("이미지가 포함된 상위 폴더를 선택하세요.");
	if (myFolder == null) {
		exit();
	}
	var myFiles = [];
	GetSubFolders(myFolder);

	// 중복 체크
	var result = false;
	if (duplicateCheck(myFiles, result)) {
		exit();
	}

	// DocuType에 따라 프로세스 분기
	var myDoc, myBook, linkResults;
	if (docuType == "activeDocument") { // 현재 열려있는 파일만
		myDoc = app.activeDocument;
		linkResults = reLinkimagePath(myDoc, myFiles);
		if (linkResults > 0) {
			alert("유실된 이미지가 있습니다. 로그 파일을 확인하세요.");
		} else alert("완료합니다.");
	} else if (docuType == "documents") { // 열려있는 파일 모두
		myDoc = app.documents;
		var sCount = 0;
		for (var a=0; a<myDoc.length; a++) {
			linkResults = reLinkimagePath(myDoc[a], myFiles);
			sCount = linkResults + sCount;
		}
		if (sCount > 0) {
			alert("유실된 이미지가 있습니다. 로그 파일을 확인하세요.");
		} else alert("완료합니다.");

	} else if (docuType == "activeBook") { // 활성화된 북 파일
		myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		for (var b=0; b<myBookContents.length; b++) {
			var myPath = myBookContents[b].fullName;
			var bookDoc = File(myPath);
			app.open(bookDoc);
		}
		var tCount = 0;
		for (var c=0; c<app.documents.length; c++) {
			linkResults = reLinkimagePath(app.documents[c], myFiles);
			tCount = linkResults + tCount;
		}
		if (tCount > 0) {
			alert("유실된 이미지가 있습니다. 로그 파일을 확인하세요.");
		} else alert("완료합니다.");
	}

	function reLinkimagePath(myDoc, myFiles) {
		// 이미지 재연결
		var doc = myDoc;
		var docPath = doc.filePath;
		var imgCount = doc.allGraphics.length;
		var linkName, newImage
		
		// 프로그레스 바
		var myProg = new Window ('palette', doc.name + ' 이미지 링크 재 연결 실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, imgCount.length);
		myProg.pbar.preferredSize.width = 300;
		myProg.show();

		// $.writeln(myFiles.length + "1111111111");
		for (var i=0; i<imgCount; i++) {
			myProg.pbar.value = i + 1;
			linkName = doc.allGraphics[i].itemLink.name;
			// $.writeln(linkName + "2222222222222");
			for (var j=0; j<myFiles.length; j++) {
				if (linkName == myFiles[j].name) {
					// $.writeln(myFiles[j].path + "/" + myFiles[j].name) + "3333333333333";
					// $.writeln("OK");
					newImage = File(myFiles[j].path + "/" + linkName);
					doc.allGraphics[i].itemLink.relink(newImage);
					// $.writeln(doc.allGraphics[i].itemLink.relink(newImage));
				}
			}
		}
		// 파일 유무 검사
		myProg.text = '파일 유무 검사 중...';
		var nLinkName;
		var errCount = 0;
		var errfilelist = [];
		for (i=0; i<imgCount; i++) {
			nLinkName = File(doc.allGraphics[i].itemLink.filePath);
			if (!nLinkName.exists) {
				errCount ++;
				errfilelist.push(doc.allGraphics[i].itemLink.name + " : " + decodeURI(nLinkName) + "\r");
			}
		}
		// 프로그레스 바 종료
		myProg.close();
		doc.save(); // 문서 저장
		
		// 유실된 파일이 있을 경우 로그를 남긴다.
		if (errCount > 0) {
			var Report = new File(docPath +"/" + doc.name.replace(".indd","") + "_log.txt");
			Report.open("w");
			Report.encoding = "utf-8";
			Report.writeln("다음의 파일들은 설정된 경로에 존재하지 않습니다.");
			for (var x = 0; x<errfilelist.length; x++) {
				Report.writeln(errfilelist[x]);
			}
			Report.close();
			return errCount;
		}
	}
	// sub functions
	function GetSubFolders(theFolder) { // 하위 폴더 모두 읽기
		var myFileList = theFolder.getFiles();
		for (var n = 0; n < myFileList.length; n++) {
			var myFile = myFileList[n];
			if (myFile instanceof Folder) {
				GetSubFolders(myFile);
			}
			else if (myFile instanceof File && myFile.name.match(/\.ai$|\.png$|\.psd$/i)) {
				myFiles.push(myFile);
			}
		}
	}

	function duplicateCheck(myFiles, result) { // 중복 체크
		for (var k=0; k<myFiles.length; k++) {
			var checkName = myFiles[k].name;
			for (var l=k+1; l<myFiles.length; l++) {
				if (checkName === myFiles[l].name) {
					result = true;
					alert(myFiles[k].name + " 파일이 중복됩니다. 중복되는 파일을 정리한 후 다시 실행하세요.");
					return result;
				}
			}
		}
	}
}