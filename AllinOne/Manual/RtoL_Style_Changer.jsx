//RtoL_Style_Changer
#targetengine "session";
function RtoL_Style_Changer() {
    if (app.documents.length == 0) {  
        alert("인디자인 문서를 열고 실행해주세요.");  
        exit();
    } else RtoL_Style_Changer_Manin();
}

function RtoL_Style_Changer_Manin() {
    var doc = app.activeDocument;
    var pStyles = doc.allParagraphStyles;
    var myRtolwin = new Window ("palette","RtoL 문서 변환기");  
    var panel_1 = myRtolwin.add ("panel");
        panel_1.add("statictext", undefined, "주의사항")
        panel_1.add("statictext", undefined, "!문서당 한번씩만 사용해주세요!")

    btn_01 = panel_1.add ("button", [0,0,190,30], "RtoL 단락/문자 스타일 변환");
    btn_08 = panel_1.add ("button", [0,0,190,30], "생가 전용 RtoL 스타일 변환");
    btn_02 = panel_1.add ("button", [0,0,190,30], "바인딩/스토리/테이블 방향 전환");
    btn_05 = panel_1.add ("button", [0,0,190,30], "선택한 테이블 RtoL 전환");
    btn_07 = panel_1.add ("button", [0,0,190,30], "선택한 테이블 LtoR 전환");
    btn_06 = panel_1.add ("button", [0,0,190,30], "고정된 개체 X 오프셋 변경");

    btn_08.onClick = function () {
        RtoL_01 ();
        RtoL_03 ();
        RtoL_02 ();
        //RtoL_05 (); //Defualt digit
        alert ("생가 전용 RtoL 스타일 변환 완료");
    }

    btn_01.onClick = function () {
        RtoL_01 ();
        RtoL_03 ();
        RtoL_02 ();
        RtoL_04 (); //arabic digit
        alert ("LtoR을 RtoL로 변환 완료");
    }
    btn_02.onClick = function () {
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
        alert ("바인딩, 스토리, 테이블의 방향 변경 완료")
    }
    btn_05.onClick = function () {
        var selTable = app.selection[0];
        selTable.tableDirection = TableDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
    }
    // btn_03.onClick = function () {
    //     var myScriptPath = MainScriptPath + list_ps_cs;
    // 	var myScriptFile = new File (myScriptPath);
    // 	$.evalFile(myScriptFile);
    // }
    // btn_04.onClick = function () {
    //     myRtolwin.close();
    //     var myScriptPath = MainScriptPath + MainScriptFile;
    // 	var myScriptFile = new File (myScriptPath);
    // 	$.evalFile(myScriptFile);
    // }
    btn_06.onClick = function () {
        var myWindow = new Window ("dialog", "X 오프셋 입력하기");
        myWindomyRtolwin.orientation = "row";
        myWindomyRtolwin.add ("statictext", undefined, "X 오프셋 입력하기(숫자 mm):");
        var myText = myWindomyRtolwin.add ('edittext {text: 0, characters: 3, justify: "center", active: true}');
        myText.characters = 20;
        myText.active = true;
        myWindomyRtolwin.add ("button", undefined, "OK");
        //var btn_cancel = myWindomyRtolwin.add ("button", undefined, "Cancel");

        
        if(myWindomyRtolwin.show() == 1) {
            if (app.documents.length > 0) {
                var setValue = Number(myText.text)*1;
                pageItems = app.documents[0].allPageItems;  
                for (i = pageItems.length-1; i >= 0; i--) {  
                    if (pageItems[i].parent instanceof Character) {  
                        pageItems[i].anchoredObjectSettings.anchorXoffset = setValue;
                    } 
                }  
            }
        }
        alert("X 오프셋 설정이 " + setValue + "mm가 적용되었습니다.");
    }
    btn_07.onClick = function () {
        var selTable = app.selection[0];
        selTable.tableDirection = TableDirectionOptions.LEFT_TO_RIGHT_DIRECTION;
    }

    myRtolwin.show();

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
}