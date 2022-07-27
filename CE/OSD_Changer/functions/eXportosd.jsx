function eXportOSD() {
	var doc = app.activeDocument;

	var file = new File(doc.filePath.absoluteURI + "/osd-list.csv"); //파일명 설정
	var ok = file.open("w");  
	if (!ok) {  
		alert("Error: " + fileToWrite.error);  
		exit();  
	}
	file.encoding = "ks_c_5601-1987";

	doc.xmlPreferences.defaultCellTagName = "Cell";
	doc.xmlPreferences.defaultImageTagName = "Image";
	doc.xmlPreferences.defaultStoryTagName = "Story";
	doc.xmlPreferences.defaultTableTagName = "Table";

	doc.deleteUnusedTags();

	var gStory = doc.stories;
	var myGraphics = doc.allGraphics;
	for (x=0; x < myGraphics.length; x++){
		try {
				myGraphics[x].autoTag();
		} catch(e) {
				// e;
				continue;
		}
	}

	var myCStyle = doc.allCharacterStyles;
	for (var i = 1 ; i < myCStyle.length ; i++) {
		var myChStyle = myCStyle[i];
		doc.xmlExportMaps.add(myChStyle, 'span');
	}

	doc.mapStylesToXMLTags();

	var rootXML = doc.xmlElements[0];
	var spanElement = rootXML.evaluateXPathExpression (".//span");
	for (var i=0; i<spanElement.length; i++) {
		var CStyle = spanElement[i].insertionPoints[0].appliedCharacterStyle;
		spanElement[i].xmlAttributes.add('style', String(CStyle.name));
	}

	//$.writeln(spanElement.length);

	var osds = rootXML.evaluateXPathExpression('//*[@style]')
	var osds_count = osds.length;
	var myPage, myID, myType, mySel
	var count = 1;
	//페이지번호 / OSD 값(순서대로) / ID / 자동반영 or 수동반영
	file.writeln ("순서,페이지" + "," + "OSD 스타일" + "," + "OSD 값" + "," + "ID" + "," + "유형" + "\r");
	for (n=0; n<osds_count; n++) {
		var osdName = osds[n].xmlAttributes.itemByName('style').value
		try {
			if (osdName == "C_OSD-NoBold" || osdName == "C_OSD") {
				myPage = osds[n].paragraphs[0].parentTextFrames[0].parentPage.name;
				myID = osds[n].textStyleRanges[0].appliedConditions[0].name;
				//app.selection[0].appliedConditions[0].indicatorColor == UIColors.GREEN
				if (osds[n].textStyleRanges[0].appliedConditions[0].indicatorColor == UIColors.GREEN) {
					myType = "수동";
				} else myType = "자동";
				//$.writeln(myPage + "\t" + osdName + "\t" + osds[n].contents + "\t" + myID + "\t" + myType);
				file.writeln (count + "," + myPage + "," + osdName + "," + osds[n].contents.replace("\r","") + "," + myID + "," + myType + "\r");
				count++;
			}
		} catch(e) {
			alert(e.line + ":" + e + "\r" + osdName + ":" + osds[n].contents);
			untagxml()
			exit();
		}
	}

	file.close();
	file.execute();
	$.sleep(500)
	untagxml();
	alert("완료합니다.");

	function untagxml() {
		var rootXML = doc.xmlElements[0];
		for (var j=rootXML.xmlElements.length-1; j>=0; j--) {
			rootXML.xmlElements[j].untag();
		}
	}
}