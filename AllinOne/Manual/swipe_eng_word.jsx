#targetengine "session";
var change = "change"
popm();

function popm() {
	var myWindow = new Window ("palette");
	var myMessage = myWindow.add ("edittext", [0,0,230,150],"", {multiline:true,readonly:true,scrolling:false});
		myMessage.text = "미주 ↔ 구주 영단어 변경 스크립트\r\r▲ 단어를 변경하면서 change 스타일을 적용합니다.\r▲ 스크립트를 실행하기 전에 문서에서 change 스타일이 적용된 문서인지 먼저 확인하기 바랍니다.\r▲ 미주 → 구주, 구주 → 미주 한번씩만 실행하시기 바랍니다.";
	var myButtonGroup = myWindow.add ("panel");
		myButtonGroup.alignment = "center";
		myButtonGroup.orientation = "column";
	btn_US2UK = myButtonGroup.add ("button", [0,0,200,30], "[Docu] 미주 → 구주 변경");
	btn_UK2US = myButtonGroup.add ("button", [0,0,200,30], "[Docu] 구주 → 미주 변경");
	btn_US2UK_book = myButtonGroup.add ("button", [0,0,200,30], "[Book] 미주 → 구주 변경");
    btn_UK2US_book = myButtonGroup.add ("button", [0,0,200,30], "[Book] 구주 → 미주 변경");
	btn_color = myButtonGroup.add ("button", [0,0,200,30], "change 스타일 색상");
	btn_none = myButtonGroup.add ("button", [0,0,200,30], "[Docu] change 스타일 삭제");
	btn_book_remove = myButtonGroup.add ("button", [0,0,200,30], "[Book] change 스타일 삭제");
	
///US2UK DOC///////////////////////////////////////////////////////
btn_US2UK.onClick = function () {
	if(app.documents.length < 1){
		alert("인디자인 문서를 여세요!");
		exit();
	} else {
		var doc = app.activeDocument;
		var cStyles = doc.allCharacterStyles;
		for (var i = 1; i < cStyles.length; i ++) {
			var cs = cStyles[i];
		}
		if (doc.characterStyles.item(change) == null) {
			alert (change + " 스타일이 없습니다. 문자 스타일을 만들고 진행합니다.");
			var new_Style = doc.characterStyles.add();
			new_Style.name = change;
			change_US2UK ();
			alert("변경완료 US to UK, change 스타일 확인 후 삭제하세요.");
		} else {
			alert (change + " 스타일이 적용된 문서입니다. change 스타일을 삭제하고 진행하세요.");
		}
	}
}
///US2UK BOOK///////////////////////////////////////////////////////
btn_US2UK_book.onClick = function () {
	if (app.books.length == 0) {
		alert("인디자인 북 파일을 열어주세요.");
	} else if (app.books.length > 0) {
		US2UK_book_opened ();
	}
	//////Book///////
	function US2UK_book_opened () {
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

		for(var n=0; n<myBookContents.length; n++) {
			var myPath = myBookContents[n].fullName;
			var myFile = File(myPath);
			var doc = app.open(myFile);
			var cStyles = doc.allCharacterStyles;
			for (var i = 1; i < cStyles.length; i ++) {
					var cs = cStyles[i];
			}
			if (doc.characterStyles.item(change) == null) {
				var new_Style = doc.characterStyles.add();
				new_Style.name = change;
				change_US2UK ();
				apply_color ();
				doc.close(SaveOptions.YES);
			} else {
				alert (change + " 스타일이 적용된 문서가 있습니다. change 스타일을 삭제하고 진행하세요.");
			}
		}
		alert("변경완료 US to UK, change 스타일 확인 후 삭제하세요.");
		// turn on warnings
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
		}
	}
	
///UK2US DOC///////////////////////////////////////////////////////
btn_UK2US.onClick = function () {
	if(app.documents.length < 1) {
		alert("인디자인 문서를 여세요!");
		exit();
	} else {
		var doc = app.activeDocument;
		var cStyles = doc.allCharacterStyles;
		for (var i = 1; i < cStyles.length; i ++) {
			var cs = cStyles[i];
		}
		if (doc.characterStyles.item(change) == null) {
			alert (change + " 스타일이 없습니다. 문자 스타일을 만들고 진행합니다.");
			var new_Style = doc.characterStyles.add();
			new_Style.name = change;
			change_UK2US ();
			alert("변경완료 UK to US, change 스타일 확인 후 삭제하세요.");
		} else {
			alert (change + " 스타일이 적용된 문서입니다. change 스타일을 삭제하고 진행하세요.");
		}
	}
}
///UK2US BOOK///////////////////////////////////////////////////////
btn_UK2US_book.onClick = function () {
	if (app.books.length == 0) {
		alert("인디자인 북 파일을 열어주세요.");
	} else if (app.books.length > 0) {
		UK2US_book_opened ();
	}
	//////Book///////
	function UK2US_book_opened () {
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

		for(var n=0; n<myBookContents.length; n++) {
			var myPath = myBookContents[n].fullName;
			var myFile = File(myPath);
			var doc = app.open(myFile);
			var cStyles = doc.allCharacterStyles;
			for (var i = 1; i < cStyles.length; i ++) {
					var cs = cStyles[i];
			}
			if (doc.characterStyles.item(change) == null) {
				var new_Style = doc.characterStyles.add();
				new_Style.name = change;
				change_UK2US ();
				apply_color ();
				doc.close(SaveOptions.YES);
			} else {
				alert (change + " 스타일이 적용된 문서가 있습니다. change 스타일을 삭제하고 진행하세요.");
			}
		}
		alert("변경완료 UK to US, change 스타일 확인 후 삭제하세요.");
		// turn on warnings
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
	}
}


btn_color.onClick = function () {
	apply_color ();
}

btn_none.onClick = function () {
	delete_color ();
}

btn_book_remove.onClick = function () {
	if (app.books.length == 0) {
		alert("인디자인 북 파일을 열어주세요.");
	} else if (app.books.length > 0) {
		delet_color_opened ();
	}
	//////Book///////
	function delet_color_opened () {
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

		for(var n=0; n<myBookContents.length; n++) {
			var myPath = myBookContents[n].fullName;
			var myFile = File(myPath);
			var doc = app.open(myFile);
			var cStyles = doc.allCharacterStyles;
			for (var i = 1; i < cStyles.length; i ++) {
					var cs = cStyles[i];
			}
			if (doc.characterStyles.item(change) == null) {
				var new_Style = doc.characterStyles.add();
				new_Style.name = change;
				change_text ();
				doc.close(SaveOptions.YES);
			} else {
				change_text ();
				doc.close(SaveOptions.YES);
			}
		}
		alert("change 스타일 삭제 완료");
		// turn on warnings
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
	}
}

myWindow.show ();
}

function delete_color () {
	if(app.documents.length < 1){
		alert ("인디자인 문서를 여세요!");
		exit();
	} else {

	var cStyles = app.activeDocument.characterStyles;
	if (cStyles.item(change) == null) {
		alert ("적용된 chnage 스타일이 없습니다.");
		exit();
	} else {
		change_text ();
		alert ("change 스타일 삭제 완료");
	}
	}
	}

function change_text () {
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
	app.findTextPreferences.findWhat = "";
	app.findTextPreferences.appliedCharacterStyle = change;
	app.changeTextPreferences.changeTo = "";
	app.changeTextPreferences.appliedCharacterStyle = app.characterStyles.item("$ID/[None]");
	//this will show changes count  
	app.activeDocument.changeText();
	//alert(app.activeDocument.changeText().length + " 개의 문자스타일 해제");
	app.activeDocument.characterStyles.item(change).remove();//chane 문자스타일 삭제
}

function apply_color () {
	var doc = app.activeDocument;
	var myColorToCheckProperties = {
		name : "check_notrans",
		colorValue : [0,100,0,0],
		space : ColorSpace.CMYK
	};
	// 색상이 있는지 없는지 확인하기
	var myColorToCheck = doc.colors.itemByName( myColorToCheckProperties.name );  
	
	if(myColorToCheck.isValid) {  
		// check_notrans 색상이 있을 경우 
		check_color ();
	};  
	
	if(!myColorToCheck.isValid) {  
		// check_notrans 색상이 없을 경우 추가하고 진행
		doc.colors.add (myColorToCheckProperties);
		check_color ();
	};

	function check_color () {
		var cs_changeName = doc.characterStyles.item(change);
		if (cs_changeName.fillColor == null) {
			cs_changeName.fillColor = "check_notrans";
		} else {
			cs_changeName.fillColor = NothingEnum.NOTHING;
		}
	}
}

function change_US2UK () {
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;

	//Set the find options.
	app.findChangeGrepOptions.includeFootnotes = false;
	app.findChangeGrepOptions.includeHiddenLayers = false;
	app.findChangeGrepOptions.includeLockedLayersForFind = false;
	app.findChangeGrepOptions.includeLockedStoriesForFind = false;
	app.findChangeGrepOptions.includeMasterPages = false;

	var greps = [
		{"findWhat":"\\<(A|a)nalog(?=.|^|$|\\n|\\r)","changeTo":"$1nalogue"},
		{"findWhat":"\\<(A|a)nalyze(?=.|^|$|\\n|\\r)","changeTo":"$1nalyse"},
		{"findWhat":"\\<(A|a)uthorize(?=.|^|$|\\n|\\r)","changeTo":"$1uthorise"},
		{"findWhat":"\\<(B|b)ehavior(?=.|^|$|\\n|\\r)","changeTo":"$1ehaviour"},
		{"findWhat":"\\<(C|c)anceled(?=.|^|$|\\n|\\r)","changeTo":"$1ancelled"},
		{"findWhat":"\\<(C|c)ategorize(?=.|^|$|\\n|\\r)","changeTo":"$1ategorise"},
		{"findWhat":"\\<(C|c)entered(?=.|^|$|\\n|\\r)","changeTo":"$1entred"},
		{"findWhat":"\\<(C|c)enter(?=.|^|$|\\n|\\r)","changeTo":"$1entre"},
		{"findWhat":"\\<(C|c)olor(?=.|^|$|\\n|\\r)","changeTo":"$1olour"},
		{"findWhat":"\\<(C|c)ustomize(?=.|^|$|\\n|\\r)","changeTo":"$1ustomise"},
		{"findWhat":"\\<(C|c)ommercialize(?=.|^|$|\\n|\\r)","changeTo":"$1ommercialise"},
		{"findWhat":"\\<(D|d)estabilize(?=.|^|$|\\n|\\r)","changeTo":"$1estabilise"},
		{"findWhat":"\\<(D|d)iscolor(?=.|^|$|\\n|\\r)","changeTo":"$1iscolour"},
		{"findWhat":"\\<(E|e)mphasize(?=.|^|$|\\n|\\r)","changeTo":"$1mphasise"},
		{"findWhat":"\\<(E|e)qualize(?=.|^|$|\\n|\\r)","changeTo":"$1qualise"},
		{"findWhat":"\\<(F|f)amiliarize(?=.|^|$|\\n|\\r)","changeTo":"$1amiliarise"},
		{"findWhat":"\\<(F|f)avorite(?=.|^|$|\\n|\\r)","changeTo":"$1avourite"},
		{"findWhat":"\\<(F|f)inalize(?=.|^|$|\\n|\\r)","changeTo":"$1inalise"},
		{"findWhat":"\\<(G|g)ray(?=.|^|$|\\n|\\r)","changeTo":"$1rey"},
		{"findWhat":"\\<(I|i)nitialize(?=.|^|$|\\n|\\r)","changeTo":"$1nitialise"},
		{"findWhat":"\\<(L|l)abeled(?=.|^|$|\\n|\\r)","changeTo":"$1abelled"},
		{"findWhat":"\\<(L|l)icense(?=.|^|$|\\n|\\r)","changeTo":"$1icence"},
		{"findWhat":"\\<(L|l)icensing(?=.|^|$|\\n|\\r)","changeTo":"$1icencing"},
		{"findWhat":"\\<(M|m)emorize(?=.|^|$|\\n|\\r)","changeTo":"$1emorise"},
		{"findWhat":"\\<(M|m)inimize(?=.|^|$|\\n|\\r)","changeTo":"$1inimise"},
		{"findWhat":"\\<(M|m)illimeter(?=.|^|$|\\n|\\r)","changeTo":"$1illimetre"},
		{"findWhat":"\\<(O|o)ptimization(?=.|^|$|\\n|\\r)","changeTo":"$1ptimisation"},
		{"findWhat":"\\<(O|o)ptimize(?=.|^|$|\\n|\\r)","changeTo":"$1ptimise"},
		{"findWhat":"\\<(O|o)ptimizing(?=.|^|$|\\n|\\r)","changeTo":"$1ptimising"},
		{"findWhat":"\\<(O|o)rganize(?=.|^|$|\\n|\\r)","changeTo":"$1rganise"},
		{"findWhat":"\\<(O|o)rganizing(?=.|^|$|\\n|\\r)","changeTo":"$1rganising"},
		{"findWhat":"\\<(P|p)ersonalize(?=.|^|$|\\n|\\r)","changeTo":"$1ersonalise"},
		{"findWhat":"\\<(P|p)olarizer(?=.|^|$|\\n|\\r)","changeTo":"$1olariser"},
		{"findWhat":"\\<(P|p)ixelization(?=.|^|$|\\n|\\r)","changeTo":"$1ixelisation"},
		{"findWhat":"\\<(P|p)rogramming(?=.|^|$|\\n|\\r)","changeTo":"$1rogramming"},
		{"findWhat":"\\<(P|p)rogram(?=.|^|$|\\n|\\r)","changeTo":"$1rogramme"},
		{"findWhat":"\\<(R|r)ecognize(?=.|^|$|\\n|\\r)","changeTo":"$1ecognise"},
		{"findWhat":"\\<(R|r)ecognizing(?=.|^|$|\\n|\\r)","changeTo":"$1ecognising"},
		{"findWhat":"\\<(S|s)tandardize(?=.|^|$|\\n|\\r)","changeTo":"$1tandardise"},
		{"findWhat":"\\<(S|s)ynchronize(?=.|^|$|\\n|\\r)","changeTo":"$1ynchronise"},
		{"findWhat":"\\<(S|s)yncronization(?=.|^|$|\\n|\\r)","changeTo":"$1yncronisation"},
		{"findWhat":"\\<(T|t)heater(?=.|^|$|\\n|\\r)","changeTo":"$1heatre"},
		{"findWhat":"\\<(U|u)nauthorize(?=.|^|$|\\n|\\r)","changeTo":"$1nauthorise"},
		{"findWhat":"\\<(U|u)tilize(?=.|^|$|\\n|\\r)","changeTo":"$1tilise"},
		{"findWhat":"\\<(V|v)isualization(?=.|^|$|\\n|\\r)","changeTo":"$1isualisation"},
	]

	for(var i=0; i < greps.length; i++) {
		var doc = app.activeDocument;
		app.findGrepPreferences.findWhat = greps[i].findWhat;
		app.findGrepPreferences.appliedCharacterStyle = app.characterStyles.item("$ID/[None]");
		app.changeGrepPreferences.appliedCharacterStyle = change;
		app.changeGrepPreferences.changeTo = greps[i].changeTo;
		doc.changeGrep();
		}
}

function change_UK2US () {
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;

	//Set the find options.
	app.findChangeGrepOptions.includeFootnotes = false;
	app.findChangeGrepOptions.includeHiddenLayers = false;
	app.findChangeGrepOptions.includeLockedLayersForFind = false;
	app.findChangeGrepOptions.includeLockedStoriesForFind = false;
	app.findChangeGrepOptions.includeMasterPages = false;

	var greps = [
		{"findWhat":"\\<(A|a)nalogue(?=.|^|$|\\n|\\r)","changeTo":"$1nalog"},
		{"findWhat":"\\<(A|a)nalyse(?=.|^|$|\\n|\\r)","changeTo":"$1nalyze"},
		{"findWhat":"\\<(A|a)uthorise(?=.|^|$|\\n|\\r)","changeTo":"$1uthorize"},
		{"findWhat":"\\<(B|b)ehaviour(?=.|^|$|\\n|\\r)","changeTo":"$1ehavior"},
		{"findWhat":"\\<(C|c)ancelled(?=.|^|$|\\n|\\r)","changeTo":"$1anceled"},
		{"findWhat":"\\<(C|c)ategorise(?=.|^|$|\\n|\\r)","changeTo":"$1ategorize"},
		{"findWhat":"\\<(C|c)entred(?=.|^|$|\\n|\\r)","changeTo":"$1entered"},
		{"findWhat":"\\<(C|c)entre(?=.|^|$|\\n|\\r)","changeTo":"$1enter"},
		{"findWhat":"\\<(C|c)olour(?=.|^|$|\\n|\\r)","changeTo":"$1olor"},
		{"findWhat":"\\<(C|c)ustomise(?=.|^|$|\\n|\\r)","changeTo":"$1ustomize"},
		{"findWhat":"\\<(C|c)ommercialise(?=.|^|$|\\n|\\r)","changeTo":"$1ommercialize"},
		{"findWhat":"\\<(D|d)estabilise(?=.|^|$|\\n|\\r)","changeTo":"$1estabilize"},
		{"findWhat":"\\<(D|d)iscolour(?=.|^|$|\\n|\\r)","changeTo":"$1iscolor"},
		{"findWhat":"\\<(E|e)mphasise(?=.|^|$|\\n|\\r)","changeTo":"$1mphasize"},
		{"findWhat":"\\<(E|e)qualise(?=.|^|$|\\n|\\r)","changeTo":"$1qualize"},
		{"findWhat":"\\<(F|f)amiliarise(?=.|^|$|\\n|\\r)","changeTo":"$1amiliarize"},
		{"findWhat":"\\<(F|f)avourite(?=.|^|$|\\n|\\r)","changeTo":"$1avorite"},
		{"findWhat":"\\<(F|f)inalise(?=.|^|$|\\n|\\r)","changeTo":"$1inalize"},
		{"findWhat":"\\<(G|g)rey(?=.|^|$|\\n|\\r)","changeTo":"$1ray"},
		{"findWhat":"\\<(I|i)nitialise(?=.|^|$|\\n|\\r)","changeTo":"$1nitialize"},
		{"findWhat":"\\<(L|l)abelled(?=.|^|$|\\n|\\r)","changeTo":"$1abeled"},
		{"findWhat":"\\<(L|l)icence(?=.|^|$|\\n|\\r)","changeTo":"$1icense"},
		{"findWhat":"\\<(L|l)icencing(?=.|^|$|\\n|\\r)","changeTo":"$1icensing"},
		{"findWhat":"\\<(M|m)emorise(?=.|^|$|\\n|\\r)","changeTo":"$1emorize"},
		{"findWhat":"\\<(M|m)inimise(?=.|^|$|\\n|\\r)","changeTo":"$1inimize"},
		{"findWhat":"\\<(M|m)illimetre(?=.|^|$|\\n|\\r)","changeTo":"$1illimeter"},
		{"findWhat":"\\<(O|o)ptimisation(?=.|^|$|\\n|\\r)","changeTo":"$1ptimization"},
		{"findWhat":"\\<(O|o)ptimise(?=.|^|$|\\n|\\r)","changeTo":"$1ptimize"},
		{"findWhat":"\\<(O|o)ptimising(?=.|^|$|\\n|\\r)","changeTo":"$1ptimizing"},
		{"findWhat":"\\<(O|o)rganise(?=.|^|$|\\n|\\r)","changeTo":"$1rganize"},
		{"findWhat":"\\<(O|o)rganising(?=.|^|$|\\n|\\r)","changeTo":"$1rganizing"},
		{"findWhat":"\\<(P|p)ersonalise(?=.|^|$|\\n|\\r)","changeTo":"$1ersonalize"},
		{"findWhat":"\\<(P|p)olariser(?=.|^|$|\\n|\\r)","changeTo":"$1olarizer"},
		{"findWhat":"\\<(P|p)ixelisation(?=.|^|$|\\n|\\r)","changeTo":"$1ixelization"},
		{"findWhat":"\\<(P|p)rogramming(?=.|^|$|\\n|\\r)","changeTo":"$1rogramming"},
		{"findWhat":"\\<(P|p)rogramme(?=.|^|$|\\n|\\r)","changeTo":"$1rogram"},
		{"findWhat":"\\<(R|r)ecognise(?=.|^|$|\\n|\\r)","changeTo":"$1ecognize"},
		{"findWhat":"\\<(R|r)ecognising(?=.|^|$|\\n|\\r)","changeTo":"$1ecognizing"},
		{"findWhat":"\\<(S|s)tandardise(?=.|^|$|\\n|\\r)","changeTo":"$1tandardize"},
		{"findWhat":"\\<(S|s)ynchronise(?=.|^|$|\\n|\\r)","changeTo":"$1ynchronize"},
		{"findWhat":"\\<(S|s)yncronisation(?=.|^|$|\\n|\\r)","changeTo":"$1yncronization"},
		{"findWhat":"\\<(T|t)heatre(?=.|^|$|\\n|\\r)","changeTo":"$1heater"},
		{"findWhat":"\\<(U|u)nauthorise(?=.|^|$|\\n|\\r)","changeTo":"$1nauthorize"},
		{"findWhat":"\\<(U|u)tilise(?=.|^|$|\\n|\\r)","changeTo":"$1tilize"},
		{"findWhat":"\\<(V|v)isualisation(?=.|^|$|\\n|\\r)","changeTo":"$1isualization"},
	]

	for(var i=0; i < greps.length; i++) {
		var doc = app.activeDocument;
		app.findGrepPreferences.findWhat = greps[i].findWhat;
		app.findGrepPreferences.appliedCharacterStyle = app.characterStyles.item("$ID/[None]");
		app.changeGrepPreferences.appliedCharacterStyle = change;
		app.changeGrepPreferences.changeTo = greps[i].changeTo;
		doc.changeGrep();
	}
}