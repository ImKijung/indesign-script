﻿function changeLang(doc, osdDB, selectLang) {
	// var doc = app.activeDocument;
	var mmiFindChnageFile = new File(osdDB);
	var osd1= doc.characterStyles.item("MMI");
	var osd2= doc.characterStyles.item("MMI_NoBold");
	var osdlist = [ osd1, osd2 ]

	//reset search
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	//set find options
	app.findChangeTextOptions.includeFootnotes = false;
	app.findChangeTextOptions.includeHiddenLayers = false;
	app.findChangeTextOptions.includeLockedLayersForFind = false;
	app.findChangeTextOptions.includeLockedStoriesForFind = false;
	app.findChangeTextOptions.includeMasterPages = false;

	//find OSD style and apply ID
	var sindex = selectLang;

	for (var i=0;i<osdlist.length;i++) {
		app.findTextPreferences.appliedCharacterStyle = osdlist[i];
		var found = app.findText();
		var myWindow = new Window ('palette');
			myWindow.add ("statictext", undefined, doc.name + " : " + osdlist[i].name + " 변환 중...");
			myWindow.pbar = myWindow.add ('progressbar', undefined, 0, found.length);
			myWindow.pbar.preferredSize.width = 300;
			myWindow.show();
		for (var j=0;j<found.length;j++) {
			myWindow.pbar.value = j+1;
			// found[j].select();
			var seltxt = found[j].contents;
			try {
				if (found[j].appliedConditions[0] == undefined) {
					// $.writeln(seltxt + "\t" + "not applied ID")
				} else {
					// $.writeln(seltxt + "\t" + app.selection[0].appliedConditions[0].name)
					var contentsID = found[j].appliedConditions[0].name;
					mmiFindChnageFile = File(mmiFindChnageFile);
					var myResult = mmiFindChnageFile.open("r", undefined, undefined);
					
					if (myResult == true) {
						do  {
							var myLine = mmiFindChnageFile.readln();
							var mmiFindChangeArray = myLine.split("\t");
							var mmiID = mmiFindChangeArray[0]; //ID
							var mmiLang = mmiFindChangeArray[sindex] //selected language
							
							if (mmiID.match(contentsID) == contentsID) {
								// applyLang(mmiID, mmiLang);
								// $.writeln(": " + mmiLang);
								if (mmiLang == "" || mmiLang == null) { // 다국어 값이 공란인 경우
									found[j].contents = "!!!" + found[j].contents;
								} else { // 다국어를 적용한다.
									found[j].contents = mmiLang;
								}
							}
						} while(mmiFindChnageFile.eof == false);
							mmiFindChnageFile.close();
					}
				}
			} catch(e) {
				$.writeln(seltxt + " - " + e.line + ":" + e);
			}
		}
		$.sleep(20); // Do something useful here
		myWindow.close();
	}
	

	//reset search
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
	// alert("OSD 언어를 변경했습니다.");
}