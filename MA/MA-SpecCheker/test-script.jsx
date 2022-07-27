var docs = app.documents;
if (docs.length == 0) {
	alert ("열려있는 인디자인 문서가 없습니다.");
} else {
	var fdoc = app.activeDocument;
	var docPath = fdoc.filePath;
	var file = new File(docPath + "/LIT-Heading-list.txt");
	file.open("w");
	file.encoding = "utf-8";
	var myWindow = new Window ('palette');
		myWindow.pbar = myWindow.add ('progressbar', undefined, 0, docs.length);
		myWindow.pbar.preferredSize.width = 150;
		myWindow.show();
	for (k=0;k<docs.length; k++) {
		myWindow.pbar.value = k+1;
		var openDocs = app.documents.everyItem().getElements();
		app.activeDocument = openDocs[openDocs.length-1];
		var doc = app.activeDocument;
		var story = doc.stories[0];
		var paras = story.paragraphs.everyItem().getElements();
		for (var i=0;i<paras.length;i++) {
			var para = paras[i];
			var pNum = para.parentTextFrames[0].parentPage.name;
			if (para.appliedParagraphStyle.name.indexOf("Heading1") != -1) {
				var tsr = para.textStyleRanges;
				for (var j=0;j<tsr.length;j++) {
					if (tsr[j].appliedCharacterStyle.name.indexOf("MMI") != -1) {
						// var pContents = para.contents.replace("\x{FEFF}","");//\x{FEFF}
						$.writeln(doc.name + " - " + pNum + " - " + para.appliedParagraphStyle.name + " - " + para.contents);
						file.writeln(doc.name + " - " + pNum + " - " + para.appliedParagraphStyle.name + " - " + para.contents);
						// $.writeln(tsr[j].contents)
					}
				}
			}
		}
		$.sleep(20); // Do something useful here
	}
	alert("LIT 헤딩 MMI 목록 추출 완료");
}
file.close();
file.execute();







