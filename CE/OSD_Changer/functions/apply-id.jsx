function applyID(osdDB, selectLang) {
	var doc = app.activeDocument;
	var mmiFindChnageFile = new File(osdDB);
	var osd1= doc.characterStyles.item("C_OSD");
	var osd2= doc.characterStyles.item("C_OSD-NoBold");
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

	var sindex = selectLang; //EN

	//find OSD style and apply ID
	try {
		for (var i=0;i<osdlist.length;i++) {
			app.findTextPreferences.appliedCharacterStyle = osdlist[i];
			var found = app.findText();
			// $.writeln(found.length);
			var myWindow = new Window ('palette');
				myWindow.pbar = myWindow.add ('progressbar', undefined, 0, found.length);
				myWindow.pbar.preferredSize.width = 300;
				myWindow.show();
			for (var j=0;j<found.length;j++) {
				myWindow.pbar.value = j+1;
				found[j].select();
				if (app.selection[0].contents == "\r") {
					app.activeWindow.activePage = found[j].parentTextFrames[0].parentPage;
					alert("문자 스타일 적용이 잘못되어 있습니다.");
					myWindow.close();
					exit();
				}
				var seltxt = app.selection[0].contents;
				findmmiID(seltxt, sindex, mmiFindChnageFile)
			}
			$.sleep(20); // Do something useful here
		}
	} catch (e) {
		alert("find 문자스타일 오류:" + e.line + ":" + e);
		myWindow.close();
		exit();
	}
	myWindow.close();

	//reset search
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;

	alert("OSD ID를 적용했습니다.");

	function findmmiID(seltxt, sindex, mmiList) {
		mmiFindChnageFile = File(mmiList);
		var myResult = mmiFindChnageFile.open("r", undefined, undefined);
		try {
			if (myResult == true) {
				do  {
					var myLine = mmiFindChnageFile.readln();
					var mmiFindChangeArray = myLine.split("\t");
					var mmiID = mmiFindChangeArray[0]; //ID
					var mmiLang = mmiFindChangeArray[sindex] //selected language
					seltxt = seltxt.replace(/\uFEFF/g, "").replace(/\u0007/g, "");
					if (seltxt == mmiLang) {
						// $.writeln(mmiID + " : " + seltxt)
						applyID(mmiID, seltxt);
						break;
					}
				} while(mmiFindChnageFile.eof == false);
					mmiFindChnageFile.close();
			}
		} catch(e) {
			alert("Find MMI 오류:" + seltxt + " " + e.line + ":" + e);
			myWindow.close();
			exit();
		}
	}

	function applyID(mmiID, seltxt) {
		if (mmiID == "") {
			alert(seltxt + "에 해당하는 MMI ID 값이 없습니다. DB txt 파일을 확인하세요.");
			myWindow.close();
			exit();
		}
		// var mmiIDx = mmiID.toString().split("_")[0]
		var doc = app.activeDocument;
		var Cond = doc.conditions;
		try {
			if (!doc.conditions.item(mmiID).isValid) {
				// $.writeln("add " + mmiIDx);
				Cond.add({
					name: mmiID,
					indicatorColor: UIColors.YELLOW,
					indicatorMethod: ConditionIndicatorMethod.useHighlight
				});
			}
			//무조건인지 조건인지 확인할 것
			//무조건이면 undefined를 출력, app.selection[0].appliedConditions.name 사용
			if (app.selection[0].appliedConditions[0] == undefined) { //무조건인 경우
				app.selection[0].appliedConditions = doc.conditions.item(mmiID);
			} else { //조건이 적용된 경우
				if (!(app.selection[0].appliedConditions[0].indicatorColor == UIColors.GREEN)) { //그린색상이 아닌 경우 조건 적용
					app.selection[0].appliedConditions = doc.conditions.item(mmiID);
				}
			}
		} catch(e) {
			alert("Apply ID 오류: " + mmiID + " / " + seltxt + " : " + e.line + ":" + e);
			myWindow.close();
			exit();
		}
	}
}