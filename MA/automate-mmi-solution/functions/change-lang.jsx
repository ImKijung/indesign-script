#include "lib.jsx";

function changeLang(doc, osdDB, selectLang) {
	var fileData = readTabDelimitedFile(osdDB);
	// $.writeln(fileData.length); //MMI ID의 갯수 -1
	
	var osd1= doc.characterStyles.item("MMI");
	var osd2= doc.characterStyles.item("MMI_NoBold");
	var osdlist = [ osd1, osd2 ]

	//reset search
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;

	//set find options
	app.findChangeGrepOptions.includeFootnotes = false;
	app.findChangeGrepOptions.includeHiddenLayers = false;
	app.findChangeGrepOptions.includeLockedLayersForFind = false;
	app.findChangeGrepOptions.includeLockedStoriesForFind = false;
	app.findChangeGrepOptions.includeMasterPages = false;

	//find OSD style and apply ID
	var sindex = selectLang;
	var seltxt, contentsID;

	for (var i=0;i<osdlist.length;i++) {
		app.findGrepPreferences.appliedCharacterStyle = osdlist[i];
		var found = doc.findGrep();
		var myWindow = new Window ('palette');
			myWindow.add ("statictext", undefined, doc.name + " : " + osdlist[i].name + " 변환 중...");
			myWindow.pbar = myWindow.add ('progressbar', undefined, 0, found.length);
			myWindow.pbar.preferredSize.width = 300;
			myWindow.show();
		var count = 0
		var cumNum;
		for (var j=0;j<found.length;j++) {
			myWindow.pbar.value = j+1;
			seltxt = found[j].contents;
			try {
				if (found[j].appliedConditions[0] == undefined) {
					// nothing
				} else {
					contentsID = found[j].appliedConditions[0].name;
					for (var n=1;n<fileData.length;n++) {
						// ID 값이 일치하는 경우
						if (fileData[n][0] == contentsID) {
							if (fileData[n][sindex] == "" || fileData[n][sindex] == null) {
								found[j].contents = "!!!" + found[j].contents;
								cumNum = j + 1;
								break;
							} else {
								found[j].contents = fileData[n][sindex];
								cumNum = j + 1;
								break;
							}
						}
					}
					if (j >= cumNum) {
						found[j].contents = "@@@" + found[j].contents;
					}
				}
			} catch(e) {
				$.writeln(seltxt + " - " + e.line + ":" + e);
			}
		}
	}
	//reset search
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
}