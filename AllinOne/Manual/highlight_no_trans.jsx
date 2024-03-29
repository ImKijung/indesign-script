#targetengine "session";
// var MainScriptPath = "C:/Program Files (x86)/Adobe/Adobe InDesign CS5.5/Scripts/Scripts Panel/";
// var MainScriptFile = "AllinOne_Script.jsx";

function highlight_no_trans_main() {
	var myhighW = new Window ("palette");
	var myMessage_1 = myhighW.add ("statictext");
		myMessage_1.text = "No trans Checker";
	var myButtonGroup = myhighW.add ("panel");
		myButtonGroup.alignment = "center";
		myButtonGroup.orientation = "column";
	btn_01 = myButtonGroup.add ("button", [0,0,200,30], "No trans 색상 적용/해제");
	btn_02 = myButtonGroup.add ("button", [0,0,200,30], "Book: No trans 색상 적용");
	btn_03 = myButtonGroup.add ("button", [0,0,200,30], "Book: No trans 색상 해제");
	// quit_btn = myButtonGroup.add ("button", [0,0,200,30], "돌아가기");


btn_01.onClick = function(){
	hightlight_notrans();
};

btn_02.onClick = function() {
	if (app.books.length == 0) {
		alert("인디자인 북 파일을 열어주세요.");
	} else if (app.books.length > 0) {
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

		for(var n=0; n<myBookContents.length; n++) {
			var myPath = myBookContents[n].fullName;
			var myFile = File(myPath);
			var doc = app.open(myFile);
			change_pink_book ();
			doc.save(myFile);
		}
		alert("Book: No trans 색상 적용");
		// turn on warnings
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
	}
}

btn_03.onClick = function() {
	if (app.books.length == 0) {
		alert("인디자인 북 파일을 열어주세요.");
	} else if (app.books.length > 0) {
		var myBook = app.activeBook;
		var myBookContents = myBook.bookContents.everyItem().getElements();
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

		for(var n=0; n<myBookContents.length; n++) {
			var myPath = myBookContents[n].fullName;
			var myFile = File(myPath);
			var doc = app.open(myFile);
			change_None_book ();
			doc.save(myFile);
		}
		alert("Book: No trans 색상 해제");
		// turn on warnings
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
	}
}

// quit_btn.onClick = function () {
// 	myhighW.close();
// 	var myScriptFile = new File (MainScriptPath + MainScriptFile);
// 	$.evalFile(myScriptFile);
// }

	myhighW.show ();
}

function hightlight_notrans () {
	if(app.documents.length < 1){
		alert ("인디자인 문서를 여세요!");
		exit();
	} else

	var doc = app.activeDocument;
	var cStyles = doc.allCharacterStyles;
	//색상 만들기
	var myColorToCheckProperties = {
		name : "check_notrans",
		colorValue : [0,100,0,0],
		space : ColorSpace.CMYK
	};
	
	// 색상이 있는지 없는지 확인하기
	var myColorToCheck = doc.colors.itemByName( myColorToCheckProperties.name );  
	
	if(myColorToCheck.isValid) {  
		// check_notrans 색상이 있을 경우 
		check_color_notrans ();
	};  
	
	if(!myColorToCheck.isValid) {  
		// check_notrans 색상이 없을 경우 추가하고 진행
		doc.colors.add (myColorToCheckProperties);
		check_color_notrans ();
	};

	// check_notrans 색상이 적용되었는지 확인
	function check_color_notrans () {
		app.findTextPreferences = NothingEnum.nothing;  
		app.changeTextPreferences = NothingEnum.nothing;  
		app.findTextPreferences.underline = true;  
		app.findTextPreferences.underlineColor = "check_notrans";  
		var myFoundStyles = doc.findText();
		var counter = 0
		try {
			for (var z = 0; z < myFoundStyles.length; z++) {
			var fStyle = myFoundStyles[z];
			counter ++;
			}
			if (counter > 0) { // check_notrans 색상이 적용됐을 경우 문자 스타일의 색상 값 삭제
				change_None ();
				alert ("하이라이트 해제");
				exit();
			}
			else (counter = 0) // check_notrans 색상이 적용 안됐을 경우 문자 스타일의 색상 값 적용
				change_pink ();
				alert ("하이라이트 적용");
				exit();
		} catch (ex) {}
	}
}
function change_pink () {
	var doc = app.activeDocument;
	var cStyles = doc.allCharacterStyles;
	var cs = doc.characterStyles;
	for (var i = 1; i < cStyles.length; i ++) {
		var cs = cStyles[i];
		switch (cs.name) {
			case "C_No-trans" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize; // 기본 단락의 크기로 두께 조정
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "C_Font-no-trans" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize; // 기본 단락의 크기로 두께 조정
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "C_No-trans-bold" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "C_Font-no-transB" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "No-trans" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "No-trans-bold" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "No-trans_comma" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "No-trans-bold-white" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "C_Notrans" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "C_Notrans-NoBold" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "No-trans_bold" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "No-trans_Red" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
			case "No-trans_Blue" : 
				cs.underline = true;
				cs.underlineWeight = doc.paragraphStyles.item("[기본 단락]").pointSize;
				cs.underlineColor = "check_notrans";
				cs.underlineOffset = "-3pt";break;
		}
	}
	app.findTextPreferences = NothingEnum.nothing;  
	app.changeTextPreferences = NothingEnum.nothing;  
	
	//alert ("하이라이트 적용")
}
function change_None () {
	var doc = app.activeDocument;
	var cStyles = doc.allCharacterStyles;
	for (var i = 1; i < cStyles.length; i ++) {
		var cs = cStyles[i];
		switch (cs.name) {
			case "C_No-trans" : cs.underline = false;break;
			case "C_Font-no-trans" : cs.underline = false;break;
			case "C_No-trans-bold" : cs.underline = false;break;
			case "C_Font-no-transB" : cs.underline = false;break;
			case "No-trans" : cs.underline = false;break;
			case "No-trans-bold" : cs.underline = false;break;
			case "No-trans_comma" : cs.underline = false;break;
			case "No-trans-bold-white" : cs.underline = false;break;
			case "C_Notrans" : cs.underline = false;break;
			case "C_Notrans-NoBold" : cs.underline = false;break;
			case "No-trans_bold" : cs.underline = false;break;
			case "No-trans_Red" : cs.underline = false;break;
			case "No-trans_Blue" : cs.underline = false;break;
			}
			}
	app.findTextPreferences = NothingEnum.nothing;  
	app.changeTextPreferences = NothingEnum.nothing;  
}

function change_pink_book () {
	var doc = app.activeDocument;
	var cStyles = doc.allCharacterStyles;
	//색상 만들기
	var myColorToCheckProperties = {
		name : "check_notrans",
		colorValue : [0,100,0,0],
		space : ColorSpace.CMYK
	};
	
	// 색상이 있는지 없는지 확인하기
	var myColorToCheck = doc.colors.itemByName( myColorToCheckProperties.name );  
	
	if(myColorToCheck.isValid) {  
		// check_notrans 색상이 있을 경우 
		change_pink ();
	};  
	
	if(!myColorToCheck.isValid) {  
		// check_notrans 색상이 없을 경우 추가하고 진행
		doc.colors.add (myColorToCheckProperties);
		change_pink ();
	};
}

function change_None_book () {
	var doc = app.activeDocument;
	var cStyles = doc.allCharacterStyles;
	//색상 만들기
	var myColorToCheckProperties = {
		name : "check_notrans",
		colorValue : [0,100,0,0],
		space : ColorSpace.CMYK
	};
	
	// 색상이 있는지 없는지 확인하기
	var myColorToCheck = doc.colors.itemByName( myColorToCheckProperties.name );  
	
	if(myColorToCheck.isValid) {  
		// check_notrans 색상이 있을 경우 
		change_None ();
	};  
	
	if(!myColorToCheck.isValid) {  
		// check_notrans 색상이 없을 경우 추가하고 진행
		doc.colors.add (myColorToCheckProperties);
		change_None ();
	};
}