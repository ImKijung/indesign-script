//RtoL_Style_Changer
#targetengine "session";

// styleMod();

function styleMod() {
	if (app.documents.length == 0) {  
			alert("인디자인 문서를 열고 실행해주세요.");  
	} else StyleModify();
}

function StyleModify() {
	var myDiagSS = new Window ("palette","단락/문자 스타일 기본 변환기");  
	var panel_1 = myDiagSS.add ("panel");
		panel_1.add("statictext", undefined, "단락스타일 변환")
	var panel_2 = myDiagSS.add ("panel");
		panel_2.add("statictext", undefined, "문자스타일 변환")

	btn_01 = panel_1.add ("button", [0,0,190,30], "커닝(Kerning): Metrics 적용");
	btn_02 = panel_1.add ("button", [0,0,190,30], "합자(Ligatures): 해제");
	btn_03 = panel_1.add ("button", [0,0,190,30], "줄바꿈(No Break): 해제");
	btn_04 = panel_1.add ("button", [0,0,190,30], "하이픈(Hyphenation): 해제");
	btn_10 = panel_1.add ("button", [0,0,190,30], "단락 레이아웃(Column): 1열 적용");
	btn_11 = panel_1.add ("button", [0,0,190,30], "자간(Tracking) 설정하기");

	btn_05 = panel_2.add ("button", [0,0,190,30], "커닝 속성 비우기(Kerning)");
	btn_06 = panel_2.add ("button", [0,0,190,30], "합자(Ligatures): 해제");
	btn_07 = panel_2.add ("button", [0,0,190,30], "줄바꿈(No Break): 해제");

	btn_01.onClick = function () {
		var doc = app.activeDocument;  
		var pstyles = doc.allParagraphStyles;
		for (var i=2; i<pstyles.length; i++){  
			pstyles[i].kerningMethod = "$ID/Metrics";
		}
		alert ("Metrics 적용 완료");
	}
	btn_02.onClick = function () {
		var doc = app.activeDocument;  
		var pstyles = doc.allParagraphStyles;
		for (var i=2; i<pstyles.length; i++){  
			pstyles[i].ligatures = false;
		}
		alert ("합자(Ligatures) 해제 완료");
	}
	btn_03.onClick = function () {
		var doc = app.activeDocument;  
		var pstyles = doc.allParagraphStyles;
		for (var i=2; i<pstyles.length; i++){  
			pstyles[i].noBreak = false;
		}
		alert ("줄바꿈(No Break) 해제 완료");
	}
	btn_04.onClick = function () {
		var doc = app.activeDocument;  
		var pstyles = doc.allParagraphStyles;
		for (var i=2; i<pstyles.length; i++){  
			pstyles[i].hyphenation = false;
		}
		alert ("하이픈(Hyphenation) 해제 완료");
	}
	btn_05.onClick = function () {
		var doc = app.activeDocument;  
		var cstyles = doc.characterStyles;
		for (var i=2; i<cstyles.length; i++){  
			cstyles[i].kerningMethod = 1851876449;
		}
		alert ("커닝 속성 비우기 완료");
	}
	btn_06.onClick = function () {
		var doc = app.activeDocument;  
		var cstyles = doc.characterStyles;
		for (var i=2; i<cstyles.length; i++){  
			cstyles[i].ligatures = false;
		}
		alert ("합자(Ligatures) 해제 완료");
	}
	btn_07.onClick = function () {
		var doc = app.activeDocument;  
		var cstyles = doc.characterStyles;
		for (var i=2; i<cstyles.length; i++){  
			cstyles[i].noBreak = false;
		}
		alert ("줄바꿈(No Break) 해제 완료");
	}
	// btn_08.onClick = function () {
	// 	var myScriptFile = new File (myScriptPath);
	// 	$.writeln(myScriptFile);
	// 	$.evalFile(myScriptFile);
	// }

	btn_10.onClick = function () {
		var doc = app.activeDocument;  
		var pstyles = doc.allParagraphStyles;
		for (var i=2; i<pstyles.length; i++){  
			pstyles[i].spanColumnType = 1163092844;
		} alert ("적용 완료");
	}

	btn_11.onClick = function () {
		var myWindow = new Window ("dialog", "자간 설정하기");
		myWindow.orientation = "row";
		myWindow.add ("statictext", undefined, "자간 값 입력하기(숫자):");
		var myText = myWindow.add ('edittext {text: 0, characters: 3, justify: "center", active: true}');
		myText.characters = 20;
		myText.active = true;
		myWindow.add ("button", undefined, "OK");
		//var btn_cancel = myWindow.add ("button", undefined, "Cancel");

		if (myWindow.show() == 1) {
			//alert("자간 설정 값" + myText.text + "완료됐습니다.");
			var setValue = Number(myText.text)*1;
			var doc = app.activeDocument;
			var pstyles = doc.allParagraphStyles;
				for (var i=2; i<pstyles.length; i++){
				pstyles[i].tracking = setValue;
				}
			alert("자간 설정 값이 " + setValue + " 적용되었습니다.");
		}
		// exit ();
	}

	myDiagSS.show();
}