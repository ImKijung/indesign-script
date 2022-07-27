var doc = app.activeDocument;
doc.xmlPreferences.defaultCellTagName = "Cell";
doc.xmlPreferences.defaultImageTagName = "Image";
doc.xmlPreferences.defaultStoryTagName = "Story";
doc.xmlPreferences.defaultTableTagName = "Table";

doc.deleteUnusedTags();

//Story에 Tag 적용하기
var myPages = doc.pages;
try {
	for(var a = 0; a < myPages.length; a++){
		var myStory = myPages[a].textFrames;
		for (var y=0; y < myStory.length; y++) {
			myStory[y].autoTag();
			}
		}
} catch(e) {}

//Lock 설정된 오브젝트 해제하기
try {
	for(var a = 0; a < myPages.length; a++){
		var item = myPages[a].allPageItems;
		for(var b = 0; b < item.length; b++){
			if(item[b].locked){
				item[b].locked = false;
			}
		}
	} 
//이미지에 Tag 적용하기
var gStory = doc.stories;
for (var y=0; y < gStory.length; y++) {
	var myGraphics = gStory[y].allGraphics;
	for (x=0; x < myGraphics.length; x++){
		myGraphics[x].autoTag();
		myGraphics[x].select();
	}
}
} catch(e) {
	alert ("그룹화된 오브젝트가 있거나 테이블의 셀 안에 테이블이 들어가 있는지 확인하세요. 문서를 저장하지 말고 닫은 후 확인하세요.");
	//exit();
}

//0번 => 단락 스타일 없슴.
//1번 => 기본 단락 스타일.
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

doc.mapStylesToXMLTags();

