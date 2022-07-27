var doc = app.activeDocument;
var myRef = doc.hyperlinks.everyItem().getElements();

// doc.xmlPreferences.defaultCellTagName = "Cell";
// doc.xmlPreferences.defaultImageTagName = "Image";
// doc.xmlPreferences.defaultStoryTagName = "Story";
// doc.xmlPreferences.defaultTableTagName = "Table";
// doc.deleteUnusedTags();

// doc.mapStylesToXMLTags();

// var item, xRef, xType, sType;
// for (var i=myRef.length-1; i>=0; i--) {
// 	if (myRef[i].source.constructor == CrossReferenceSource) {
// 		item = myRef[i].source.sourceText;
// 		xType = myRef[i].source.appliedFormat.name;
// 		if (xType == "페이지 번호") {
// 			sType = "pageNum";
// 		} else if (xType == "단락 텍스트") {
// 			sType = "paraText";
// 		} else if (xType == "텍스트 앵커 이름") {
// 			sType = "textAnchor";
// 		} else {
// 			sType = "else";
// 		}
// 		// $.writeln(myRef[i].destination.parent.name + "#" + myRef[i].destination.name);
// 		xRef = doc.xmlElements[0].xmlElements.add('xref', item);
// 		xRef.xmlAttributes.add('href', String(myRef[i].destination.parent.name + "#" + myRef[i].destination.name));
// 		xRef.xmlAttributes.add('type', String(sType));
// 	}
// }
for (i=myRef.length-1; i>=0; i--) {
	if (myRef[i].source.constructor == CrossReferenceSource) {
		for (i=myRef.length-1; i>=0; i--) {
			myRef[i].source.sourceText.remove();
		}
	}
}
// var cRfs = doc.xmlElements[0].evaluateXPathExpression('//xref');
// var cRfType, cRfFormat, cRfSource, target, trDoc, sText, mySource, myDest;
// for (var j=0; j<cRfs.length; j++) {
// 	cRfType = cRfs[j].xmlAttributes.itemByName('type').value;
// 	if (cRfType == "pageNum") {
// 		cRfFormat = doc.crossReferenceFormats.itemByName("페이지 번호");
// 	} else if (cRfType == "paraText") {
// 		cRfFormat = doc.crossReferenceFormats.itemByName("단락 텍스트");
// 	} else if (cRfType == "textAnchor") {
// 		cRfFormat = doc.crossReferenceFormats.itemByName("텍스트 앵커 이름");
// 	} else {
// 		cRfFormat = doc.crossReferenceFormats.itemByName("전체 단락");
// 	}
// 	cRfSource = cRfs[j].xmlAttributes.itemByName('href').value;
// 	target = cRfSource.split("#");
// 	var rvName = "TM_23MY_ARA";
// 	if (target[0].indexOf("INDEX") != -1) {
// 		trDocName = target[0].replace(/(\d{3})_([^>]+)_(INDEX)_(\w{3})(.indd)/g, '$1' + '_' + rvName.replace(rvName.slice(-4), '') + '_' + '$3' + '_' + '$4' +'.indd');
// 		//011_TM_22MY_INDEX_RUS.indd
// 	} else if (target[0].indexOf("Appendix") != -1) {
// 		trDocName = target[0].replace(/(\d{3})_([^>]+)_(Appendix)_(\w{3})(.indd)/g, '$1' + '_' + rvName.replace(rvName.slice(-4), '') + '_' + '$3' + '_' + '$4' +'.indd');
// 	} else {
// 		trDocName = target[0].replace(/(\d{3})_([^>]+)(.indd)/g, '$1' + '_' + rvName + '.indd');
// 	}
// 	$.writeln(target[0] + "#" + target[1]);
// 	// trDoc = app.documents.itemByName(trDocName);
// 	trDoc = app.documents.itemByName(target[0]);
// 	myDest = trDoc.hyperlinkTextDestinations.itemByName(target[1]);
// 	// cRfs[j].contents = "11";
// 	mySource = doc.crossReferenceSources.add(cRfs[j].texts[0].insertionPoints[-1], cRfFormat);
// 	doc.hyperlinks.add({
// 		source: mySource,
// 		destination: myDest,
// 		highlight : HyperlinkAppearanceHighlight.NONE
// 	})
// }
// var root = doc.xmlElements[0];

// for (var k=0; k<root.xmlElements.length; k++) {
// 	root.xmlElements[k].untag();
// }