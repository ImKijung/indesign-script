function _findEmptyTalbe(docuType) {
	var docCount;
	if (docuType == "activeDocument") {
		docCount = 1;
		findEmptyTalbes(docCount);
	} else if (docuType == "documents") {
		docCount = app.documents.length;
		findEmptyTalbes(docCount);
	} else if (docuType == "activeBook") {
		// 북파일
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		for (var b=0; b<myBookContents.length; b++) {
			var myPath = myBookContents[b].fullName;
			var myDoc = File(myPath);
			app.open(myDoc);
		}
		docCount = app.documents.length;
		findEmptyTalbes(docCount);
	}
}


function findEmptyTalbes(count) {
	// $.writeln(count);
	var docs = app.documents;

	var emptyStyles = [ "Empty_1", "Empty_2", "Empty_3" ];
	var zoomSet = 400;

	var myProg = new Window ('palette', 'Empty 표 검증');
		myProg.pbar = myProg.add ('progressbar', undefined, 0, count);
	myProg.pbar.preferredSize.width = 300;
	myProg.show();

	for (var x=0; x<count; x++) {
		// $.writeln(docs[x].name);
		myProg.pbar.value = x + 1;
		myProg.text = docs[x].name + ' 검사 중...';

		var paraStyles = docs[x].allParagraphStyles;
		for (var n=0; n<emptyStyles.length; n++) {
			noEmptyTable(docs[x], emptyStyles[n]);
			checkTableStyle(docs[x]);
			// $.writeln(emptyStyles[n] +" : table style OK");
		}
		var Empty2 = [ "Step-Cmd-OL1", "Step-Cmd-OL2", "Step-Description_2", "Step-UL1_2", "Step-UL1_2-Note", "UL1-Caution-Warning", "UL1_1-Note", "Empty_2", "Step-UL2_2-Note", "UL1_1", "Step-UL2_2-Note", "Description_1-Center" ];
		var Empty3 = "Step-UL1_2";

		var pages = docs[x].pages;
		var tf, paras, pStyle, prevStyle;

		for (var i=0; i<pages.length; i++) {
			var tf = pages[i].textFrames;
			for (var j=0; j<tf.length; j++) {
				var paras = tf[j].paragraphs;
				for (var k=0; k<paras.length; k++) {
					curPara = paras[k];
					pStyle = paras[k].appliedParagraphStyle.name;
					if (k > 0) {
						prevStyle = paras[k-1].appliedParagraphStyle.name;
					}
					if (pStyle == "Empty_2") {
						if (chkMatchStyle2(prevStyle)) {
							app.activeDocument = docs[x];
							curPara.select();
							app.activeWindow.zoomPercentage = zoomSet;
							myProg.close();
							alert(prevStyle + "+" + pStyle + " 스타일이 잘못 적용되어 있습니다.");
							exit();
						}
					} else if (pStyle == "Empty_3") {
						if (chkMatchStyle3(prevStyle)) {
							app.activeDocument = docs[x];
							curPara.select();
							app.activeWindow.zoomPercentage = zoomSet;
							myProg.close();
							alert(prevStyle + "+" + pStyle + " 스타일이 잘못 적용되어 있습니다.");
							exit();
						}
					}
				}
			}
		}
	}
	myProg.close();
	alert("완료합니다.");

	function chkMatchStyle2(pStyle) {
		var count = 0;
		for (var i=0;i<Empty2.length;i++) {
			if (Empty2[i] == pStyle ) {
				count ++;
			}
		}
		if (count == 0) {
			return true;
		} else return false
	}

	function chkMatchStyle3(pStyle) {
		var count = 0;
		for (var i=0;i<Empty3.length;i++) {
			if (Empty3[i] != pStyle ) {
				count ++;
			}
		}
		if (count == 0) {
			return true;
		} else return false
	}

	function noEmptyTable(doc, emptyStyle) { // Empty가 테이블인지 확인
		// app.activeDocument = doc;
		for (var i=1; i<paraStyles.length; i++) {
			if (paraStyles[i].name == emptyStyle) {
				findStyle = paraStyles[i];
				break
			}
		}

		app.findTextPreferences = NothingEnum.nothing;  
		app.changeTextPreferences = NothingEnum.nothing;
		
		app.findTextPreferences.appliedParagraphStyle = findStyle;
		
		var findEmpty = doc.findText();
		
		for (var j=0;j<findEmpty.length;j++) {
			// var charCode = findEmpty[j].paragraphs[0].contents.charCodeAt(0);
			if (findEmpty[j].paragraphs[0].contents.charCodeAt(0) != 22) {
				myProg.close();
				app.activeDocument = doc;
				findEmpty[j].select();
				app.activeWindow.zoomPercentage = zoomSet;
				alert("표가 아닌 위치에 Empty 스타일이 적용됐습니다. 수정하세요.");
				app.findTextPreferences = NothingEnum.nothing;  
				app.changeTextPreferences = NothingEnum.nothing;
				exit();
			}
		}
		
		app.findTextPreferences = NothingEnum.nothing;  
		app.changeTextPreferences = NothingEnum.nothing;
	}

	function checkTableStyle(doc) { // Table에 Empty가 아닌 경우 찾기
		// app.activeDocument = doc;
		var tables = doc.stories.everyItem().tables.everyItem().getElements();
		var tableStyle;
		for (var y=0; y<tables.length; y++) {
			tableStyle = tables[y].storyOffset.appliedParagraphStyle.name;
			// $.writeln((y+1) + ":" + tableStyle);
			if (tableStyle.indexOf("Empty") != -1) {
				// OK
			} else {
				myProg.close();
				app.activeDocument = doc;
				tables[y].storyOffset.select();
				app.activeWindow.zoomPercentage = zoomSet;
				alert(doc.name + " : 표 항목에 Empty가 아닌 문장 스타일이 적용돼 있습니다.");
				exit();
			}
		}
	}
}