function RtoL_01 () {
    var doc = app.activeDocument;
    var pStyles = doc.allParagraphStyles;
    //alert ("폰트 기본 변경은 SamsungOneArabic입니다.");
    for (var i=1; i<pStyles.length; i++) {
        var ps = pStyles[i];
        
        //문장 스타일 폰트 변경
        ps.basedOn = "[No Paragraph Style]";
        /*switch (ps.appliedFont.name) {  
        case "SamsungOne\t200": ps.appliedFont = "SamsungOneArabic\t400"; break;  
        case "SamsungOne\t400": ps.appliedFont = "SamsungOneArabic\t400"; break;  
        case "SamsungOne\t500": ps.appliedFont = "SamsungOneArabic\t450"; break;  
        case "SamsungOne\t600": ps.appliedFont = "SamsungOneArabic\t600"; break;  
        case "SamsungOne\t700": ps.appliedFont = "SamsungOneArabic\t600"; break;  
        }*/
        ps.appliedLanguage = "아랍어"; //언어 설정 변경
        ps.composer = "Adobe World-Ready 단락 컴포저"; //컴포저 변경
    }
    }
// 좌우 인덴트 변경하기
function RtoL_02 () {
    var doc = app.activeDocument;
    var pStyles = doc.allParagraphStyles;
    var cStyle, cRightIndent, fLindent;
    // jump over [None]  
    pStyles.shift();  
    // iterate through all  
    while (cStyle = pStyles.shift()) {
        cRightIndent = cStyle.rightIndent;
        cStyle.rightIndent = cStyle.leftIndent;  
        cStyle.leftIndent = cRightIndent;
        pStyles.firstLineIndent = fLindent;
    }
}
function RtoL_03 () {
    var doc = app.activeDocument;
    var pStyles = doc.allParagraphStyles;
    for (var k=1; k<pStyles.length; k++) {
        var ps = pStyles[k];
        ps.bulletsAlignment = 1919379572; // Bullet, Right 정렬
        ps.numberingAlignment = 1919379572; // 넘버링 Right 정렬
        
        ps.paragraphJustification = 1886019954; //Arabic justification
        ps.paragraphDirection = 1379028068; //Left to Right
        //정렬 변경
        switch (ps.justification) {
            case 1818915700: ps.justification = 1919578996; break; //Left justification to Right justification
            case 1818584692: ps.justification = 1919379572; break; //Left align to Right align
            case 1919379572: ps.justification = 1818584692; break; //Right align to Left align
        }

        //스타일 명에 따른 예외 사항 별도 적용
        switch (ps.name) {
            case "Description-Licence" : ps.justification = 1818584692; break; //left
            case "Description-Cell-Right" : ps.justification = 1818584692; break; //left
            case "Img-Left" : ps.justification = 1818584692; break; //left
            case "Img-Right" : ps.justification = 1919379572; break; //right
            case "Img-Center" : ps.justification = 1667591796; break; //center
        }

        //Chapter 스타일의 경우, 단락 경계선의 좌우 인덴트값 변경하기
        switch (ps.name) {
            var pAboveLindent;
            case "Chapter" : 
                pAboveLindent = ps.ruleAboveLeftIndent;
                ps.ruleAboveLeftIndent = ps.ruleAboveRightIndent; 
                ps.ruleAboveRightIndent = pAboveLindent;
        }
        
    }
    //폰트 속성 변경하기
    for (var h=1; doc.allCharacterStyles.length>h; h++){
        var cs = doc.allCharacterStyles[h]
        if(cs.fontStyle == "400")
            cs.fontStyle = "400";
        else if (cs.fontStyle == "200")
            cs.fontStyle = "400";
        else if (cs.fontStyle == "500")
            cs.fontStyle = "450";
        else if (cs.fontStyle == "600")
            cs.fontStyle = "600";
        else if (cs.fontStyle == "700")
            cs.fontStyle = "600";
    }
}
function RtoL_04 () {
    var doc = app.activeDocument;
    var pStyles = doc.allParagraphStyles;
    for (var q=1; q<pStyles.length; q++) {
        var ps = pStyles[q];
        ps.digitsType = 1684627826; // Arabic digits
    }
}

function RtoL_05 () {
    var doc = app.activeDocument;
    var pStyles = doc.allParagraphStyles;
    for (var q=1; q<pStyles.length; q++) {
        var ps = pStyles[q];
        ps.digitsType = 1684628581; // Default digits
    }
}

function directionTransform() {
    var doc = app.activeDocument;
    doc.documentPreferences.pageBinding = PageBindingOptions.RIGHT_TO_LEFT;
    for (var s = 0 ; s < doc.stories.length ; s++) {
        var myStory = doc.stories.item(s);
        myStory.storyPreferences.storyDirection = StoryDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
    }
    var a = doc.stories.everyItem().tables.everyItem().getElements();
    for (var i=0; i<a.length; i++){
        a[i].tableDirection = TableDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
    }
    var b = doc.stories.everyItem().tables.everyItem().cells.everyItem().tables.everyItem().getElements();
    for (var j=0; j<b.length; j++){
        b[j].tableDirection = TableDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
    }
}

function fixedObjectXoffset() {
    var myWindow = new Window ("dialog", "X 오프셋 입력하기");
	myWindow.orientation = "row";
	myWindow.add ("statictext", undefined, "X 오프셋 입력하기(숫자 mm):");
	var myText = myWindow.add ('edittext {text: 0, characters: 3, justify: "center", active: true}');
	myText.characters = 20;
	myText.active = true;
	myWindow.add ("button", undefined, "OK");
	//var btn_cancel = myWindow.add ("button", undefined, "Cancel");

	if(myWindow.show() == 1) {
		if (app.documents.length > 0) {
			var setValue = Number(myText.text)*1;
			pageItems = app.documents[0].allPageItems;
			for (i = pageItems.length-1; i >= 0; i--) {
				if (pageItems[i].parent instanceof Character) {
					pageItems[i].anchoredObjectSettings.anchorXoffset = setValue;
				}
			}
		}
	alert("X 오프셋 설정이 " + setValue + "mm가 적용되었습니다.");
	} else {}
	
}

function moveTextFrame() {
    var myPages = app.activeDocument.pages;
    var pageNumber, allTF, j, x2, leftPostion, pagewidth;
    var lastPageNum = myPages.lastItem().name;

    for (var i=0; i<myPages.length; i++) {
        pageNumber = myPages[i].name.split("-")[1];

        try {
            if (pageNumber % 2 == 0) {
                if (lastPageNum == myPages[i].name) {
                    pagewidth = 0; // 마지막 페이지일 경우, 페이지 너비 값이 0
                } else {
                    pagewidth = 148; // 마지막 페이지가 아닌 경우 한 페이지의 너비를 뺀다.
                }
                // $.writeln(myPages[i].name + ": 짝수");
                // 짝수인 경우 geometricBounds의 x2 값에서 페이지 너비(148mm)를 빼고 [ex. 159-148 = 11mm]
                // 안쪽 페이지 너비는 17mm, 이동해야 하는 길이는 [11-17 = -6mm]
                allTF = myPages[i].textFrames;
                for (j=0; j<allTF.length; j++) {
                    x2 = allTF[j].geometricBounds[1];
                    leftPostion = x2 - pagewidth - 17;
                    // $.writeln(myPages[i].name + ": 짝수 " + leftPostion);
                    allTF[j].move(undefined, [-leftPostion, 0]);
                }
            } else {
                // $.writeln(myPages[i].name + ": 홀수");
                // 홀수인 경우 geometricBounds의 x2 값에서 [ex. 17mm]
                // 바깥쪽 페이지 너비는 9mm, 이동해야 하는 길이는 [9 - 17 = -8mm]
                allTF = myPages[i].textFrames;
                for (j=0; j<allTF.length; j++) {
                    x2 = allTF[j].geometricBounds[1];
                    leftPostion = x2 - 9;
                    // $.writeln(myPages[i].name + ": 홀수 " + leftPostion);
                    allTF[j].move(undefined, [-leftPostion, 0]);
                }
            }
        } catch(e) {
            alert(e.line + ":" + e.message);
        }
    }
}