function batcheXportpdf() {
	var myBook = app.activeBook;
	var myBookContents = myBook.bookContents.everyItem().getElements();

	// 경고창 비활성화
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	var selectedDocus = [];
	for (var b=0; b<myBookContents.length; b++) {
		selectedDocus.push(myBookContents[b].name);
	}

	var pdfPreset = app.pdfExportPresets.everyItem().getElements();

	var myDlg = new Window("dialog", "파일 별 PDF 내보내기");
		myDlg.add("statictext", undefined, "인디자인 문서 선택");
	var fileGroup = myDlg.add("group");
		fileGroup.alignChildren = "left";
		fileGroup.minimumSize.width = 160;
		var fileList = fileGroup.add("listbox", undefined, selectedDocus, {multiselect:true});
	var presetGroup = myDlg.add("group");
		presetGroup.alignChildren = "left";
		var presetlist = presetGroup.add("dropdownlist", [0, 0, 160, 30]);
	for (var j=0; j<pdfPreset.length; j++) {
		presetlist.add("item", pdfPreset[j].name);
	}
	myDlg.add("button", undefined, "OK");

	var result = myDlg.show();

	if (result == 1) { // OK 버튼
		if (presetlist.selection == null || fileList.selection == null) {
			alert("파일 또는 PDF 프리셋을 선택하세요.");
			exit();
		}
		var seletedfiles = fileList.selection;
		var presetName = presetlist.selection.text;
		exportPDFdocument(seletedfiles, presetName);
	} else {
		exit();
	}

	function exportPDFdocument(sfiles, presetName) {
		// $.writeln(sfiles.length + ":" + presetName);
		var targetFolder = Folder.selectDialog( "PDF를 저장할 폴더를 선택하세요." );
		if (targetFolder == null) {
			exit();
		}
		var sbook = app.activeBook;
		var sbookContents = sbook.bookContents.everyItem().getElements();
		
		for (var n=0; n<sbookContents.length; n++) {
			if (sfiles.join("|").indexOf(sbookContents[n].name) != -1) {
				var fullPath = myBookContents[n].fullName;
				var sDoc = File(fullPath);
				app.open(sDoc, true);
			}
		}
		var reOpenbook = app.activeBook.filePath + "/" + app.activeBook.name;
		// 메모리 점유율을 낮추기 위해 열려있는 북 파일을 닫는다.
		app.activeBook.close(SaveOptions.YES);

		var docs = app.documents;
		app.pdfExportPreferences.viewPDF = false; // PDF 출력 후 열기 옵션 끔
		var doCount = [];
		for (n=0; n<docs.length; n++) {
			if (sfiles.join("|").indexOf(docs[n].name) != -1) {
				doCount.push(n);
				// 이미지 유실 확인하기
				for (var d = docs[n].links.length-1; d >= 0; d--) {  
					var link = docs[n].links[d];  
					if (link.status == LinkStatus.LINK_MISSING) {  
						alert("현재 문서에 유실된 이미지가 있습니다. 이미지를 다시 연결한 다음 진행하세요.");
						app.open(File(reOpenbook));
						app.pdfExportPreferences.viewPDF = true; // PDF 출력 후 열기 옵션 켬
						exit();
					}
				}
				var targefile = new File(targetFolder + "/" + docs[n].name.replace(/\.indd$/, ".pdf"));
				var t = docs[n].asynchronousExportFile(ExportFormat.PDF_TYPE, targefile, false, app.pdfExportPresets.item(presetName));
				t.waitForTask();
			}
		}
		var f = doCount[0];
		// $.writeln(docs[f].name);
		t = docs[f].asynchronousExportFile(ExportFormat.PDF_TYPE, new File(targetFolder + "/" + docs[f].name.replace(/\.indd$/, ".pdf")), false, app.pdfExportPresets.item(presetName));
		t.waitForTask();

		// 문서 닫기
		// for (n = docs.length-1; n >= 0; n--) {
		// 	docs[n].close(SaveOptions.YES);
		// }

		alert("완료합니다.");
		// 북 파일을 다시 연다.
		app.open(File(reOpenbook));
		app.pdfExportPreferences.viewPDF = true; // PDF 출력 후 열기 옵션 켬
	}

	// 경고창 활성화
	// app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}