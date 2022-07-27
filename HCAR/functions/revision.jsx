function _reVision() {
	var myBook = app.activeBook;

	var rvName = prompt("변경하려는 북 파일 이름을 입력하세요.", myBook.name.replace('.indb', ''), "리비전 입력");

	if (rvName == null) {
		exit();
	}

	// 모든 파일을 열고 상호참조 XML 구조 생성
	for (var x=0; x<app.activeBook.bookContents.length; x++) {
		myDoc = app.open(app.activeBook.bookContents[x].fullName);
	}

	var myWindow = new Window ('palette', "리비전 파일 생성 중 ...");
		myWindow.pbar = myWindow.add ('progressbar', undefined, 0, app.documents.length);
		myWindow.pbar.preferredSize.width = 300;
	myWindow.show();

	var myDoc = app.documents;
	for (x=0; x<myDoc.length; x++) {
		myWindow.pbar.value = x + 1;
		generateHyperlinksXml(myDoc[x]);
	}
	for (x = myDoc.length-1; x >= 0; x--) {
		myDoc[x].close(SaveOptions.YES);
	}

	myWindow.pbar.value = 0;

	var bookPath = myBook.filePath;
	var myPath = myBook.fullName;
	var rvBookName = rvName + ".indb";
	myBook.close(SaveOptions.NO);
	myPath.rename(rvBookName);

	app.open(File(bookPath + "/" + rvBookName));

	var rvBook = app.activeBook;
	var BookContents = rvBook.bookContents.everyItem().getElements();
	var document, docName, rvDocName, addRvDoc;
	var nDocs = [];
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	myWindow.pbar.text = "문서 이름 변경 중 ...";

	try {
		for (var i=BookContents.length-1;i>=0;i--) {
			// $.writeln(rvBook.bookContents[i].fullName);
			document = rvBook.bookContents[i].fullName;
			docName = rvBook.bookContents[i].name;
			if (docName.indexOf("INDEX") != -1) {
				rvDocName = docName.replace(/(\d{3})_([^>]+)_(INDEX)_(\w{3})(.indd)/g, '$1' + '_' + rvName.replace(rvName.slice(-4), '') + '_' + '$3' + '_' + '$4' +'.indd');
				//011_TM_22MY_INDEX_RUS.indd
			} else if (docName.indexOf("Appendix") != -1) {
				rvDocName = docName.replace(/(\d{3})_([^>]+)_(Appendix)_(\w{3})(.indd)/g, '$1' + '_' + rvName.replace(rvName.slice(-4), '') + '_' + '$3' + '_' + '$4' +'.indd');
			} else {
				rvDocName = docName.replace(/(\d{3})_([^>]+)(.indd)/g, '$1' + '_' + rvName + '.indd');
			}
			document.rename(rvDocName);
			addRvDoc = File(bookPath + "/" + rvDocName);
			nDocs.push(bookPath + "/" + rvDocName);
			rvBook.bookContents[i].remove();
		}
	} catch(e) {
		alert(e.line + ":" + e);
		exit();
	}

	nDocs.sort();

	// $.writeln(nDocs);
	var addDoc;

	myWindow.pbar.text = "변경한 문서 추가 중 ...";
	myWindow.pbar.value = 0;

	for (var j=0;j<nDocs.length;j++) {
		myWindow.pbar.value = j + 1;
		addDoc = new File(nDocs[j]);
		app.activeBook.bookContents.add(addDoc);
	}

	var cBook = app.activeBook;
	var cBookContents = cBook.bookContents.everyItem().getElements();
	var docNums = cBookContents.length;
	for (var k=0;k<docNums;k++) {
		var cDoc = cBookContents[k].fullName;
		app.open(cDoc);
	}

	myWindow.pbar.text = "상호 참조 재 연결 중 ...";
	myWindow.pbar.value = 0;

	for (k=0;k<docNums;k++) {
		myWindow.pbar.value = k + 1;
		var doc = app.documents[k];
		relinkCrossreference(doc, rvName);
		doc.save();
	}
	cBook.save();
	myWindow.close();
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
}

function relinkCrossreference(doc, rvName) {
	var myRef = doc.hyperlinks.everyItem().getElements();
	if (myRef.length > 1) {
		try {
			for (i=myRef.length-1; i>=0; i--) {
				if (myRef[i].source.constructor == CrossReferenceSource) {
					myRef[i].source.sourceText.remove();
				}
			}
		} catch(e) {
			alert(e.line + ":" + e);
			exit();
		}

		var cRfs = doc.xmlElements[0].evaluateXPathExpression('//xref');
		if (cRfs.length > 0) {
			var cRfType, cRfFormat, cRfSource, target, trDoc, mySource, myDest;
			try {
				for (var j=0; j<cRfs.length; j++) {
					cRfType = cRfs[j].xmlAttributes.itemByName('type').value;
					if (cRfType == "pageNum") {
						cRfFormat = doc.crossReferenceFormats.itemByName("페이지 번호");
					} else if (cRfType == "paraText") {
						cRfFormat = doc.crossReferenceFormats.itemByName("단락 텍스트");
					} else if (cRfType == "textAnchor") {
						cRfFormat = doc.crossReferenceFormats.itemByName("텍스트 앵커 이름");
					} else {
						cRfFormat = doc.crossReferenceFormats.itemByName("전체 단락");
					}
					cRfSource = cRfs[j].xmlAttributes.itemByName('href').value;
					target = cRfSource.split("#");
					
					if (target[0].indexOf("INDEX") != -1) {
						trDocName = target[0].replace(/(\d{3})_([^>]+)_(INDEX)_(\w{3})(.indd)/g, '$1' + '_' + rvName.replace(rvName.slice(-4), '') + '_' + '$3' + '_' + '$4' +'.indd');
						//011_TM_22MY_INDEX_RUS.indd
					} else if (target[0].indexOf("Appendix") != -1) {
						trDocName = target[0].replace(/(\d{3})_([^>]+)_(Appendix)_(\w{3})(.indd)/g, '$1' + '_' + rvName.replace(rvName.slice(-4), '') + '_' + '$3' + '_' + '$4' +'.indd');
					} else {
						trDocName = target[0].replace(/(\d{3})_([^>]+)(.indd)/g, '$1' + '_' + rvName + '.indd');
					}
					// $.writeln(trDocName + " : " + target[1]);
					trDoc = app.documents.itemByName(trDocName);
					myDest = trDoc.hyperlinkTextDestinations.itemByName(target[1]);
					// cRfs[j].contents = "11";
					mySource = doc.crossReferenceSources.add(cRfs[j].texts[0].insertionPoints[-1], cRfFormat);
					doc.hyperlinks.add({
						source: mySource,
						destination: myDest,
						highlight : HyperlinkAppearanceHighlight.NONE
					})
				}
			} catch(e) {
				alert(e.line + ":" + e);
				exit();
			}
			var root = doc.xmlElements[0];
			for (var k=0; k<root.xmlElements.length; k++) {
				root.xmlElements[k].untag();
			}
		}
	}
}

function generateHyperlinksXml(doc) {
	var myRef = doc.hyperlinks.everyItem().getElements();
	var item, xRef, xType, sType;
	
	if (myRef.length > 1) {
		doc.xmlPreferences.defaultCellTagName = "Cell";
		doc.xmlPreferences.defaultImageTagName = "Image";
		doc.xmlPreferences.defaultStoryTagName = "Story";
		doc.xmlPreferences.defaultTableTagName = "Table";
		doc.deleteUnusedTags();
		doc.mapStylesToXMLTags();

		try {
			for (var i=myRef.length-1; i>=0; i--) {
				if (myRef[i].source.constructor == CrossReferenceSource) {
					item = myRef[i].source.sourceText;
					xType = myRef[i].source.appliedFormat.name;
					if (xType == "페이지 번호") {
						sType = "pageNum";
					} else if (xType == "단락 텍스트") {
						sType = "paraText";
					} else if (xType == "텍스트 앵커 이름") {
						sType = "textAnchor";
					} else {
						sType = "else";
					}
					// $.writeln(myRef[i].destination.parent.name + "#" + myRef[i].destination.name);
					xRef = doc.xmlElements[0].xmlElements.add('xref', item);
					xRef.xmlAttributes.add('href', String(myRef[i].destination.parent.name + "#" + myRef[i].destination.name));
					xRef.xmlAttributes.add('type', String(sType));
				}
			}
		} catch(e) {
			alert(e.line + ":" + e);
			exit();
		}
	}
}