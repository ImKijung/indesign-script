var doc = app.activeDocument;

doc.xmlPreferences.defaultCellTagName = "Cell";
doc.xmlPreferences.defaultImageTagName = "Image";
doc.xmlPreferences.defaultStoryTagName = "Story";
doc.xmlPreferences.defaultTableTagName = "Table";

var myPStyle = doc.allParagraphStyles;
var myTag = null;
for (var i = 2 ; i < myPStyle.length ; i++) {
	var myPaStyle = myPStyle[i];// 단락 스타일
	doc.xmlExportMaps.add(myPaStyle, myPaStyle.name);
}

//0번 => 문자 스타일 없슴.
var myCStyle = doc.allCharacterStyles;
for (var i = 1 ; i < myCStyle.length ; i++) {
	var myChStyle = myCStyle[i];
	doc.xmlExportMaps.add(myChStyle, myChStyle.name);
}

//0번 => 표 스타일 없슴.
//1번 => 기본 표 스타일.
var allTableStyles = doc.allTableStyles;
for (var i = 2 ; i < allTableStyles.length ; i++) {
	var TbStyle = allTableStyles[i];
	doc.xmlExportMaps.add(TbStyle, TbStyle.name);
}

//0번 => 셀 스타일 없슴.
var allCellStyles = doc.allCellStyles;
for (var i = 1 ; i < allCellStyles.length ; i++) {
	var cellStyle = allCellStyles[i];
	doc.xmlExportMaps.add(cellStyle, cellStyle.name);
}

doc.mapStylesToXMLTags();