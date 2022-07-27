function applyID(doc) {
	// var doc = app.activeDocument;
	var cstyles = doc.allCharacterStyles;
	var mmiBold = doc.characterStyles.item("MMI");
	var mmiNoBold = doc.characterStyles.item("MMI_NoBold");
	var mmiArray = [ mmiBold, mmiNoBold ];
	var cdns = doc.conditions;
	var MMI = cdns.item("mmi-unassign");

	if (MMI.isValid == true) {
		// $.writeln("있다.");
	} else {
		// $.writeln("없다.");
		// mmi-unassign 조건이 없으니 추가한다.
		cdns.add({
			name: "mmi-unassign",
			indicatorColor: [255, 255, 0],
			indicatorMethod: ConditionIndicatorMethod.useHighlight
		});
	}

	//set find options
	app.findChangeTextOptions.includeFootnotes = false;
	app.findChangeTextOptions.includeHiddenLayers = false;
	app.findChangeTextOptions.includeLockedLayersForFind = false;
	app.findChangeTextOptions.includeLockedStoriesForFind = false;
	app.findChangeTextOptions.includeMasterPages = false;

	appliyConditions();
	applyIDformmi();
	//완료

	function appliyConditions() {
		for (var j=0; j<mmiArray.length; j++ ) {
			
			app.findTextPreferences = NothingEnum.nothing;
			app.changeGrepPreferences = NothingEnum.nothing;

			app.findTextPreferences.appliedCharacterStyle = mmiArray[j];
			app.changeTextPreferences.appliedConditions = ["mmi-unassign"];

			doc.changeText();
		}
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
	}

	function applyIDformmi() {
		//apply-ID-모델명.txt 불러오기
		var count
		// var doc = app.activeDocument;
		var mmiList = doc.filePath.absoluteURI + "/" + doc.name.split(".indd")[0] + ".txt"; //파일명 설정
		if (mmiList == null) {
			alert("ID를 적용할 기준 txt 파일이 없습니다.");
			exit();
		}
		var mmiFindChnageFile = File(mmiList);
		var myResult = mmiFindChnageFile.open("r", undefined, undefined);

		if (myResult == false) {
			alert(mmiList + "\r파일을 찾을 수 없습니다.");
			doc.conditions.everyItem().remove();
			exit();
		}
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing
		app.findTextPreferences.appliedConditions = ["mmi-unassign"];
		var foundmmi = doc.findText();
		// $.writeln(foundmmi.length);
		if (doc.name.indexOf("Settings") != -1) {// Settings 챕터 제목은 0974 아이디로 수동으로 적용한다.
			count = 2;
			if (!doc.conditions.item("MMI-0199").isValid) {
				// $.writeln("add " + mmiIDx);
				doc.conditions.add({
					name: "MMI-0199",
					indicatorColor: UIColors.GRID_GREEN,
					indicatorMethod: ConditionIndicatorMethod.useHighlight
				});
			}
			foundmmi[0].appliedConditions = doc.conditions.item("MMI-0199");
		} else
			count = 1;
		
		do {
			var myLine = mmiFindChnageFile.readln();
			var mmiFindChangeArray = myLine.split("\t");
			var mmiID = mmiFindChangeArray[0]; //ID
			var mmiLang = mmiFindChangeArray[1] //String
			try {
				//mmiID 값을 conditions에 추가
				if (!doc.conditions.item(mmiID).isValid) {
					// $.writeln("add " + mmiIDx);
					doc.conditions.add({
						name: mmiID,
						indicatorColor: UIColors.GRID_GREEN,
						indicatorMethod: ConditionIndicatorMethod.useHighlight
					});
				}
				if (mmiLang != foundmmi[count-1].contents.replace(/\ufeff/g, "")) {
					alert(count + "번째 ID의 값(" + mmiLang + ")과 문서의 MMI 값(" + foundmmi[count-1].contents + ")이 서로 다릅니다.");
					exit();
				}
				foundmmi[count-1].appliedConditions = doc.conditions.item(mmiID);
				// $.writeln(count + " - " + mmiID + " : " + mmiLang + "\t" + foundmmi[count-1].contents);
				count ++;
			} catch(e) {
				alert(e + ":" + "mmi-unassign 조건이 적용되지 않았습니다.");
			}
		} while(mmiFindChnageFile.eof == false);
			mmiFindChnageFile.close();
			
		app.findTextPreferences = NothingEnum.nothing;
		app.changeTextPreferences = NothingEnum.nothing;
	}
}