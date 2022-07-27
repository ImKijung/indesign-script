// Written by Kasyan Servetsky
// August 30, 2011
// http://www.kasyan.ho.com.ua
// e-mail: askoldich@yahoo.com
// List of all paragraph styles and character styles, showing the font, font size and leading.
// 단락 스타일과 문자 스타일의 목록을 출력하는 스크립드
//        InDesign으로 연 현재 문서에 정의된
//        단락 스타일과 문자 스타일의 목록을 다음의 기본 속성들과 함께 csv 파일로 출력한다.
//        단락 스타일의 경우:
//                name 스타일 이름
//                appliedFont.postscriptName  폰트명
//                pointSize 폰트 크기
//                leading  레딩, 즉 행간
///        문자 스타일의 경우:
//                name 스타일 이름
//                appliedFont  폰트명
//                fontStyle 폰스 스타일 (italic이 적용되었는지, bold가 적용되었는지 등)
//                pointSize 폰트 크기
//                leading  레딩, 즉 행간
//===============================================================

//======================= FUNCTIONS  ============================
function List_pscs() {
	var gDoc = app.activeDocument;
	var ps, cs;
	var txt = "Paragraph styles:\r#,Style name,Base Style,Font name,Font style,Font size,Leading,Align(정렬),First Line Indent,Left Indent,Right Indent,Tracking(자간),Color,Horizontal Scale,Vertical Scale,Baseline Shift,Kerning,language,Ligatures(합자),No Break,Hyphenination,Span Column,불렛스타일,넘버링스타일\r"  // Delimeter를  세미콜론(;)에서 콤마(,)로 변경함 , Excel에서 열이 나누어져 열리게  하기 위함
                                                                                                                       // csv는 comma separated value의 약자임

    //  문서의 모든 단락 스타일을 담고 있는 컬렉션인 allParagraphStyles을 이용해서 프로세싱한다.
    var parStyles = gDoc.allParagraphStyles;

	for (var p = 1; p < parStyles.length; p++) {
		ps = parStyles[p];
		txt = txt + (p) + "," + ps.name + "," + ((ps.basedOn.name == null) ? "Not set" : ps.basedOn.name) + "," + ps.appliedFont.postscriptName + "," + ps.fontStyle + "," + ps.pointSize + "," + ((ps.leading == 1635019116) ? "Auto" : ps.leading) + "," + ((ps.justification == 1818584692 ) ? "Left" : ps.justification) + "," + ps.firstLineIndent + "," + ps.leftIndent + "," + ps.rightIndent + "," + ((ps.tracking == 1635019116) ? "Auto" : ps.tracking) + "," + ((ps.fillColor.name == "") ? "Unnamed color" : ps.fillColor.name) + "," + ps.horizontalScale + "," + ps.verticalScale + "," + ps.baselineShift + "," + ps.kerningMethod + "," + ps.appliedLanguage.name + "," + ps.ligatures + "," + ps.noBreak + "," + ps.hyphenation + "," + ((ps.spanColumnType == 1163092844) ? "Single Column" : ps.spanColumnType) + "," + ps.bulletsCharacterStyle.name + "," + ps.numberingCharacterStyle.name + "," + "\r";
	}

    //  문서의 모든 문자 스타일을 담고 있는 컬렉션인 allCharacterStyles을 이용해서 프로세싱한다.
	var charStyles = gDoc.allCharacterStyles;
	txt = txt + "\rCharacter styles:\r#,Style name,Base Style,Font name,Font style,Font size,Leading,Tracking(자간),Color,Horizontal Scale,Vertical Scale,Baseline Shift,Kerning,language,Ligatures(합자),No Break\r";
	for (var c = 1; c < charStyles.length; c++) {
		cs = charStyles[c];
		txt = txt + (c) + ","+ cs.name + ","+ ((cs.basedOn.name == null) ? "Not set" : cs.basedOn.name) + "," + ((cs.appliedFont == "") ? "Not set" : cs.appliedFont)  + ","+ ((cs.fontStyle ==1851876449 ) ? "Not set" : cs.fontStyle)+ ","+ ((cs.pointSize == 1851876449) ? "Not set" : cs.pointSize) + ","+ ((cs.leading == 1851876449) ? "Not set" : cs.leading) + ","+ ((cs.tracking == 1851876449) ? "Not set" : cs.tracking) + "," + ((cs.fillColor == null) ? "Unnamed color" : cs.fillColor.name) + "," + ((cs.horizontalScale == 1851876449) ? "Not set" : cs.horizontalScale) + "," + ((cs.verticalScale == 1851876449) ? "Not set" : cs.verticalScale) + "," + ((cs.baselineShift == 1851876449) ? "Not set" : cs.baselineShift) + "," + ((cs.kerningMethod == 1851876449) ? "Not set" : cs.kerningMethod) + "," + ((cs.appliedLanguage == null) ? "Not set" : cs.appliedLanguage.name) + "," + ((cs.ligatures == 1851876449) ? "Not set" : cs.ligatures) + "," + ((cs.noBreak == 1851876449) ? "Not set" : cs.noBreak) + ","+ "\r";
	}

    //  바탕화면에 내용(txt)이 기록된 파일을 만들고, 이것을 기본 연결 프로그램(.csv 파일의 경우, 엑셀)로 연다. 
	WriteToFile(txt);
}

//--------------------------------------------------------------------------------------------------------------
// 새 빈 파일을 만든 후,  내용을 쓰고,  기본 연결 프로그램(default associated application)으로  그 파일을 여는 함수
function WriteToFile(text) {
	file = new File("~/Desktop/Paragraph and character styles.csv"); // 파일은 바탕화면("~/Desktop/")에 만들어진다.
	// file.encoding = "UTF-8";
    file.encoding = "ks_c_5601-1987"; // 파일 내용에 대한 설정 부분: 인코딩 값을 "ks_c_5601-1987"로 지정하면 엑셀에서 한글이 깨지지 않고 나타난다.
	file.open("w");  // 파일의 내용에 대해 작업하기 위해서는 가장 먼저 열어야 한다.
	file.write(text); 
	file.close(); // 파일의 내용에 대해 작업한 후에는 항상 닫아야 한다.
	file.execute(); // 작성한 .csv 파일을 default associated application으로 연다. (.csv 파일의  default associated application은 MS Excel이다.)
}
//--------------------------------------------------------------------------------------------------------------

function ErrorExit(error, icon) {  
	alert(error, gScriptName + " - " + gScriptVersion, icon);
	exit();
}
//--------------------------------------------------------------------------------------------------------------