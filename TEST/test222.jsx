var docs = app.documents;
	for (var y=0;y<docs.length;y++){
		relinkCrossreference(docs[y]);
	}
	// alert ("모든 상호 참조를 재연결했습니다.");

	// relinkCrossreference();
	function relinkCrossreference(doc) {
		// var doc = app.activeDocument;
		var myCross = doc.crossReferenceSources.everyItem().getElements();
		var myRef = doc.hyperlinks;
		for (var c=0; c<myRef.length;c++) {
			try {
				var dest = myRef[c].destination;
				if (dest instanceof HyperlinkTextDestination) {
					// app.select(myRef[i].source.sourceText);
					// var mySourceText = app.selection[0];
					var mySourceText = myRef[c].source.sourceText;
					var targetDocument = myRef[c].destination.parent.name;
					var myDestiny = myRef[c].destination.name;
					// $.writeln(i+1 + " " + mySourceText.contents + " - " + targetDocument + " - " + myDestiny);
					myRef[c].remove();
					mySourceText.remove();

					var xRefFormat = doc.crossReferenceFormats.itemByName("전체 단락"); 
					var mySource = doc.crossReferenceSources.add(mySourceText, xRefFormat);
					// var myHypDest = doc.hyperlinkTextDestinations.itemByName("Navigation bar (soft buttons)");
					// var mySource = doc.crossReferenceSources.item(i);
					var target = app.documents.itemByName(targetDocument);
					var myHypDest = target.hyperlinkTextDestinations.itemByName(myDestiny);
					// var highlight = HyperlinkAppearanceHighlight.INVERT;
					doc.hyperlinks.add(
						{ source: mySource, 
						destination : myHypDest,
						highlight : HyperlinkAppearanceHighlight.INVERT
						}
					);
				}
				else continue;
			} catch(e) {
				alert(e);
			}
		}
		doc.save();
	}