	var doc = app.activeDocument;

	doc.xmlPreferences.defaultCellTagName = "Cell";
	doc.xmlPreferences.defaultImageTagName = "Image";
	doc.xmlPreferences.defaultStoryTagName = "Story";
	doc.xmlPreferences.defaultTableTagName = "Table";

	doc.deleteUnusedTags();


	//Hyperlink의 Tag 적용하기
	var myRef = doc.crossReferenceSources.everyItem().getElements();
	for (var i=0; i<myRef.length; i++) {
		app.select(myRef[i].sourceText);
		selection = app.selection[0];
		doc.xmlElements[0].xmlElements.add('xref', selection).xmlAttributes.add('href', String(myRef[i].name));
	}

	//Destination의 Tag 적용하기
	link = doc.hyperlinkTextDestinations;
	for (var i=0 ; i<link.length; i++) {
		var linkName = link[i].name;
		var xxx = link[i].destinationText;
		xxx.select();
		selection = app.selection[0];
		doc.xmlElements[0].xmlElements.add('anchor', selection).xmlAttributes.add('x', String(linkName));
	}
