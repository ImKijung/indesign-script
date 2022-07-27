function eXportAddmmi(newID) {
	var root = new XML("<root/>");
	var BMs = new XML("<items/>");
	root.appendChild(BMs);
	
	var doc = app.activeDocument;
	var osd1= doc.characterStyles.item("MMI");
	var osd2= doc.characterStyles.item("MMI_NoBold");
	var osdlist = [ osd1, osd2 ]
	var mmiID = [];
	var cDoc, docuName
	//reset search
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	//set find options
	app.findChangeTextOptions.includeFootnotes = false;
	app.findChangeTextOptions.includeHiddenLayers = false;
	app.findChangeTextOptions.includeLockedLayersForFind = false;
	app.findChangeTextOptions.includeLockedStoriesForFind = false;
	app.findChangeTextOptions.includeMasterPages = false;

	for (var i=0;i<osdlist.length;i++) {
		app.findTextPreferences.appliedCharacterStyle = osdlist[i];
		var found = app.findText();

		for (var j=0; j<found.length; j++) {
			try {
				if (found[j].appliedConditions[0] == undefined) {
					//$.writeln("MMI-" + fillZero(4, newID + "") + " : " + found[j].contents);
					mmiID = "MMI-" + fillZero(4, newID + "");
					// var newID = fillZero(4, lLastrow + "");
					if (found[j].parent.parent.constructor.name == "Table") {
						cDoc = found[j].parent.parent.parent.parent.parent;
					} else {
						cDoc = found[j].parent.parent;
					}
					if (!cDoc.conditions.item(mmiID).isValid) {
						// $.writeln("add " + mmiIDx);
						cDoc.conditions.add({
							name: mmiID,
							indicatorColor: UIColors.BRICK_RED,
							indicatorMethod: ConditionIndicatorMethod.useHighlight
						});
					}
					docuName = cDoc.name.replace(/\d\d\d\_/g, "").replace(".indd","");
					found[j].appliedConditions = cDoc.conditions.item(mmiID);
					item = <item id={mmiID} docu={docuName}>{found[j].contents}</item>;
					BMs.appendChild(item);
					newID++;
				}
			} catch(e) {
				alert(e.line + ":" + e);
				exit();
			}
		}
	}

	//reset search
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	// var file = new File(bookPath + "/" + book.name.split(".indb").join(".xml"));
	var file = new File(doc.filePath.absoluteURI + "/newID.xml"); //파일명 설정
	var xml = root.toXMLString();
	file.open("w");
	file.encoding = "UTF-8";
	file.write(xml);
	
	file.close();
}

function fillZero(width, str) {
	return str.length >= width ? str:new Array(width-str.length+1).join('0')+str;//남는 길이만큼 0으로 채움
}