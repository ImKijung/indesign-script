function batchExportIDML() {
	var myFolder = Folder.selectDialog("인디자인 문서가 있는 상위 폴더를 선택하세요.");
	if (myFolder == null) {
		exit();
	}
	var targetFolder = Folder.selectDialog("IDML 파일을 저장할 폴더를 선택하세요.");
	if (targetFolder == null) {
		exit();
	}

	if (myFolder != null && targetFolder != null) {
		var myFiles = [];
		GetSubFolders(myFolder);
		if (myFiles.length > 0) {
			app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
			// 프로그레서 바 출현
			var myProg = new Window ('palette', '실행 중...');
			myProg.pbar = myProg.add ('progressbar', undefined, 0, myFiles.length);
			myProg.pbar.preferredSize.width = 300;
			myProg.show();
			
			for (var i=0; i<myFiles.length; i++) {
				myProg.text = myFiles[i].name + ' 변환 중...';
				myProg.pbar.value = i + 1;
				if (myFiles[i].name.indexOf("003_TOC") != -1 || myFiles[i].name.indexOf("016_Index") != -1) {
					// 003_Toc, 016_Index 파일은 idml을 출력하지 않는다.
					continue;
				} else {
					var myDoc = app.open(File(myFiles[i]), false);
					eXportIDML(myDoc, targetFolder);
				}
			}
		}
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
		myProg.close();
		alert("완료합니다.");
	} else exit();

	function eXportIDML(myDoc, targetFolder) {
		myDoc.save();
		var myfPath = targetFolder + "/" + myDoc.name.split(".indd")[0] + ".idml";
		var myFile = new File(myfPath);
		myDoc.exportFile(ExportFormat.INDESIGN_MARKUP, myFile);
		myDoc.close(SaveOptions.NO);
	}

	function GetSubFolders(theFolder) { // 하위 폴더 모두 읽기
		var myFileList = theFolder.getFiles();
		for (var x = 0; x < myFileList.length; x++) {
			 var myFile = myFileList[x];
			 if (myFile instanceof Folder) {
				  GetSubFolders(myFile);
			 }
			 else if (myFile instanceof File && myFile.name.match(/\.indd$/i)) {
				  myFiles.push(myFile);
			 }
		}
	}
}

