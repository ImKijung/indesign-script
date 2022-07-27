function changeSemicaseDown () {
	var doc = app.documents;
	for (var j=0; j< doc.length; j++) {
		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
		app.findGrepPreferences.findWhat = "(\\;) ([A-Z])|(\\;) ([\\x{0400}-\\x{04F0}])";
		var myFound = doc[j].findGrep();
		var counter = 0;
		for (var i=0; i < myFound.length; i++) {
			myFound[i].changecase(ChangecaseMode.LOWERCASE);
			counter ++;
		} alert (doc[j].name + "문서에서 " + counter + " 개의 문장을 수정했습니다.")
		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	} 
	for (var j=0;j<doc.length;j++) {
		//GIF, Bluetooth, SIM, USIM, PIN, Wi-Fi 되돌려 놓을 것
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;

		//Set the find options.
		app.findChangeGrepOptions.includeFootnotes = false;
		app.findChangeGrepOptions.includeHiddenLayers = false;
		app.findChangeGrepOptions.includeLockedLayersForFind = false;
		app.findChangeGrepOptions.includeLockedStoriesForFind = false;
		app.findChangeGrepOptions.includeMasterPages = false;

		var greps = [
			{"findWhat":"(\\;) gIF","changeTo":"$1 GIF"},
			{"findWhat":"(\\;) bluetooth","changeTo":"$1 Bluetooth"},
			{"findWhat":"(\\;) sIM","changeTo":"$1 SIM"},
			{"findWhat":"(\\;) uSIM","changeTo":"$1 USIM"},
			{"findWhat":"(\\;) pIN","changeTo":"$1 PIN"},
			{"findWhat":"(\\;) wi-Fi","changeTo":"$1 Wi-Fi"},
			{"findWhat":"(\\;) aR Doodle","changeTo":"$1 AR Doodle"},
			{"findWhat":"(\\;) always On Display","changeTo":"$1 Always On Display"},
			{"findWhat":"(\\;) samsung Members","changeTo":"$1 Samsung Members"},
			{"findWhat":"(\\;) galaxy Wearable","changeTo":"$1 Galaxy Wearable"},
			{"findWhat":"(\\;) game Launcher","changeTo":"$1 Game Launcher"},
			{"findWhat":"(\\;) samsung Pay","changeTo":"$1 Samsung Pay"},
			{"findWhat":"(\\;) samsung Daily","changeTo":"$1 Samsung Daily"},
			{"findWhat":"(\\;) samsung Notes","changeTo":"$1 Samsung Notes"},
			{"findWhat":"(\\;) uSB","changeTo":"$1 USB"},
			{"findWhat":"(\\;) kids","changeTo":"$1 Kids"},
			{"findWhat":"(\\;) toit","changeTo":"$1 Toit"},
			{"findWhat":"(\\;) game Launcher","changeTo":"$1 Game Launcher"},
			{"findWhat":"(\\;) qR","changeTo":"$1 QR"},
		]
		for(var m=0; m<greps.length; m++) {
			app.findGrepPreferences.findWhat = greps[m].findWhat;
			app.changeGrepPreferences.changeTo = greps[m].changeTo;
			doc[j].changeGrep();
		}
		app.findGrepPreferences = NothingEnum.nothing;
		app.changeGrepPreferences = NothingEnum.nothing;
	}
}