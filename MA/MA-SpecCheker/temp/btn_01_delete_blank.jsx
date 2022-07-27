if(app.documents.length < 1){
	alert("인디자인 문서를 여세요!");
	exit();
	}else{
	main ();
	}

function main() {
	var doc = app.activeDocument;
	var pStyleGroup = app.activeDocument.paragraphStyleGroups.itemByName('Contents');
	var blank = pStyleGroup.paragraphStyles.itemByName('blank_page')
	app.findTextPreferences = NothingEnum.nothing;
	app.changeTextPreferences = NothingEnum.nothing;
	
	//블랭크 페이지 찾기
	app.findTextPreferences.findWhat = "";
	app.findTextPreferences.appliedParagraphStyle = blank;
	var cf_spec = doc.findText();
	var counter = 0

	try {
		for (var z = 0; z < cf_spec.length; z++) {
		counter ++;
		}
		if (counter > 0) {
			for (var i = 0; i < cf_spec.length; i++) {
			var myPage = cf_spec[i].parentTextFrames[0].parentPage;
				myPage.remove();
			} alert ("Blank 페이지 삭제 완료");
	}
	else (counter = 0) 
		alert ("Blank 페이지가 없습니다.");
		exit();
	} catch (ex) {
		exit();
		}
}