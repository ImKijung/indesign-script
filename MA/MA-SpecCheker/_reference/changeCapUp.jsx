function changecaseUp () {
	var doc = app.documents;
	for (var j=0; j< doc.length; j++) {
		app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
		app.findGrepPreferences.findWhat = "(\\:) ([a-z])";
		var myFound = doc[j].findGrep();
		var counter = 0;
		for (var i=0; i < myFound.length; i++) {
			myFound[i].changecase(ChangecaseMode.UPPERCASE);
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
			{"findWhat":"(\\:) gIF","changeTo":"$1 GIF"},
			{"findWhat":"(\\:) bluetooth","changeTo":"$1 Bluetooth"},
			{"findWhat":"(\\:) sIM","changeTo":"$1 SIM"},
			{"findWhat":"(\\:) uSIM","changeTo":"$1 USIM"},
			{"findWhat":"(\\:) pIN","changeTo":"$1 PIN"},
			{"findWhat":"(\\:) wi-Fi","changeTo":"$1 Wi-Fi"},
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