function unused_style_delete() {
	if (app.documents.length == 0 && app.books.length == 0) {
		alert("인디자인 또는 북파일을 여세요.");
	} else if (app.books.length > 0 && app.documents.length == 0) {
		unUsedStyleBook();
		alert ("사용하지 않는 스타일을 모두 정리했습니다. 모든 문서의 처음으로 이동해 임의로 만들어진 텍스트 프레임(넘버링, 불렛 문자 스타일)을 삭제하세요.")
	} else if (app.documents.length > 0 && app.books.length == 0) {
		unUsedStyleDocu();
		alert ("사용하지 않는 스타일을 모두 정리했습니다. 문서의 처음으로 이동해 임의로 만들어진 텍스트 프레임(넘버링, 불렛 문자 스타일)을 삭제하세요.")
	} else if (app.documents.length > 0 && app.books.length > 0) {
		alert("인디자인 또는 북파일 중 하나를 닫아주세요.");
	}
}

function unUsedStyleBook() {
	var myBook = app.activeBook;
	var myBookContents = myBook.bookContents.everyItem().getElements();
	// turn off warnings: missing fonts, links, etc.
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

	for(var n=0; n<myBookContents.length; n++) {
		var myPath = myBookContents[n].fullName;
		var myFile = File(myPath);
		var doc = app.open(myFile);
		unUsedStyleDocu(doc);
		doc.save(myFile);
		//doc.close(SaveOptions.YES);
	}
}

function unUsedStyleDocu() {
	var myDoc = app.activeDocument; 
	var myParStyles = myDoc.allParagraphStyles; 
	var myCharStyles = myDoc.allCharacterStyles;
	for (var i = myDoc.allParagraphStyles.length - 1; i >= 2; i--) { 
		for (var z = myDoc.allParagraphStyles.length-1; z>=2; z--) { 
			var goodToRemove = true; 
			if (myDoc.allParagraphStyles[z].basedOn == myDoc.allParagraphStyles[i]) { 
				goodToRemove = false; 
				break; 
			} 
		} 
		if (goodToRemove) { 
			removeUnusedParaStyle1(myParStyles[i]); 
		} 
	}
	removeUnusedCharStyle();
	

	function removeUnusedParaStyle1(myPaStyle) {
		app.findTextPreferences = NothingEnum.nothing; 
		app.changeTextPreferences = NothingEnum.nothing; 
		app.findTextPreferences.appliedParagraphStyle = myPaStyle;
		app.findChangeTextOptions.includeMasterPages = true; // 마스터페이지 포함 
		var myFoundStyles = myDoc.findText(); 
			if (myFoundStyles == 0) {
				myPaStyle.remove(); 
			}
		app.findTextPreferences = NothingEnum.nothing; 
		app.changeTextPreferences = NothingEnum.nothing; 
		}

	function removeUnusedCharStyle() { 
		var gDoc = app.activeDocument;
		var paraStyles = gDoc.allParagraphStyles;
		var myCharStyles = gDoc.allCharacterStyles;
		for (j = 1; j < paraStyles.length; j++) {
			var ps = paraStyles[j];
				if (ps.bulletsCharacterStyle.name !== "[없음]") {
					var textFrame = gDoc.pages[0].textFrames.add();  
						textFrame.properties =  
						{  
							geometricBounds : [ 0,0,10,40 ],  
							strokeWidth : 0,  
							fillColor : "None",  
							contents : ps.bulletsCharacterStyle.name
						};
						textFrame.paragraphs.item(0).appliedParagraphStyle = gDoc.paragraphStyles.item("[기본 단락]");
						textFrame.characters.everyItem().appliedCharacterStyle = gDoc.characterStyles.item(ps.bulletsCharacterStyle.name);
				}
				else if (ps.numberingCharacterStyle.name !== "[없음]") {
					var textFrame = gDoc.pages[0].textFrames.add();  
						textFrame.properties =  
						{  
							geometricBounds : [ 0,0,10,40 ],  
							strokeWidth : 0,  
							fillColor : "None",  
							contents : ps.bulletsCharacterStyle.name
						};
						textFrame.paragraphs.item(0).appliedParagraphStyle = gDoc.paragraphStyles.item("[기본 단락]")
						textFrame.characters.everyItem().appliedCharacterStyle = gDoc.characterStyles.item(ps.numberingCharacterStyle.name);
				}
		}
		for (var o = gDoc.characterStyles.length - 1; o >= 2; o--) { 
					app.findTextPreferences = NothingEnum.nothing; 
					app.changeTextPreferences = NothingEnum.nothing; 
					app.findTextPreferences.appliedCharacterStyle = myCharStyles[o]; 
					var myFoundStyles = gDoc.findText(); 
						if (myFoundStyles == 0) { 
							myCharStyles[o].remove(); 
					} 
					app.findTextPreferences = NothingEnum.nothing; 
					app.changeTextPreferences = NothingEnum.nothing;
			} 
	}
}