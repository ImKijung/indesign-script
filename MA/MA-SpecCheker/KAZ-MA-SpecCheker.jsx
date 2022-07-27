#targetengine "session";
#include "_reference/RemoveBlankPages.jsx";

var w = new Window ('palette {text: "KAZ 사양 Check"}');
	//w.main = w.add ('group');
	group = w.add ('group {orientation: "column"}');
	group.orientation = "row";
		btn_02 = group.add("button", [0, 0, 150, 30], "Blank Page 삭제");
		btn_01 = group.add("button", [0, 0, 150, 30], "배터리 문구 확인");
	group02 = w.add ('group {orientation: "column"}');
	group02.orientation = "row";
		btn_03 = group02.add("button", [0, 0, 150, 30], "커버 EAC 로고 적용");
		btn_04 = group02.add("button", [0, 0, 150, 30], "-");

//----------------------------------------------------------------------------------------------------
btn_01.onClick = function specChecker01() {
	if(app.documents.length < 1){
		alert("인디자인 문서를 여세요!");
		exit();
		} else {
		var doc = app.documents;
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		for (var i=0; i < doc.length; i++){
			//app.activeDocument = doc[i];
			app.findTextPreferences.findWhat = "Батареяны алып тастау";
			app.findTextPreferences.appliedParagraphStyle = "Heading1";
			var finds = doc[i].findText();
			
			if (finds.length > 0) { // something has been found
				alert("배터리 분리 방법 설명글을 다음의 페이지에서 찾았습니다. 페이지: " + finds[0].parentTextFrames[0].parentPage.name + "\r배터리 분리 방법 설명 글을 삭제해주세요.");
				app.activeDocument = doc[i];
				app.activeWindow.activePage = finds[0].parentTextFrames[0].parentPage;
				app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
				exit();
			}
			else {
				continue;
			}
		}
		alert("배터리 분리 방법 설명글을 찾을 수 없습니다.");
		app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
		}
}
//----------------------------------------------------------------------------------------------------
btn_02.onClick = function () {
	RemoveBlankPages();
}
//----------------------------------------------------------------------------------------------------
btn_03.onClick = function () {
	var doc = app.documents;
	var dest_file = "001_Cover.indd";
	for (var i=0; i < doc.length; i++){
		if (doc[i].name == dest_file) {
			app.activeDocument = doc[i];
			//alert ("찾았다!");
			var img = app.activeDocument.pages.item(0).place(new File("C:/MA-WikiSpec/KAZ/M-eac.ai"));
			var imgFrame = img[0].parent;
			imgFrame.move(["98.75mm", "240mm"]);
			getImage()
		} else {
			continue;
		}
	} alert (dest_file + " 문서가 없습니다.");

	function getImage() {
		var mDoc = app.activeDocument;
		var mLinks = mDoc.links.everyItem().getElements();
		//var destFolder = mDoc.filePath.parent + "/image/";
		var myFolder = Folder.selectDialog ('현재 문서의 이미지가 저장된 폴더를 선택하세요. 불러온 이미지를 이동합니다.');
		var len = mLinks.length;
		if ( !Folder(myFolder).create() ) exit();

		while (len-->0) {
			var currLinkFName = mLinks[len].filePath;
			var currFile = File(myFolder + "/" + File(currLinkFName).name);
			if (!currFile.exists && File(currLinkFName).exists) {
				mLinks[len].copyLink(currFile);
				} else
			mLinks[len].relink(currFile);
		}
		alert ("이미지가 선택한 폴더로 이동되었습니다.");
		exit();
	}
}
//----------------------------------------------------------------------------------------------------
w.show();