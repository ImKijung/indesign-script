var docs = app.documents;
var chIcon;
for (var x=0; x<docs.length; x++) {
	if (docs[x].name == "001_Cover.indd" || docs[x].name == "002_TOC.indd" || docs[x].name == "007_Copyright.indd" || docs[x].name == "008_Copyright.indd") {
		continue
	} else {
		if (docs[x].name.indexOf("Basics") != -1) {
			chIcon = "I-chapter_basics";
		} else if (docs[x].name.indexOf("Applications") != -1) {
			chIcon = "I-chapter_apps";
			var chDesc = docs[x].pages[0].textFrames[0].paragraphs
			for (var u=0;u<chDesc.length;u++) {
				if (chDesc[u].appliedParagraphStyle.name == "Description") {
					chDesc[u].appliedParagraphStyle = docs[x].paragraphStyles.itemByName("Description-Chapter");
				}
				chDesc[u].clearOverrides(OverrideType.ALL);
			}
		} else if (docs[x].name.indexOf("Settings") != -1) {
			chIcon = "I-chapter_settings";
			var chDesc = docs[x].pages[0].textFrames[0].paragraphs
			for (var u=0;u<chDesc.length;u++) {
				if (chDesc[u].appliedParagraphStyle.name == "Description") {
					chDesc[u].appliedParagraphStyle = docs[x].paragraphStyles.itemByName("Description-Chapter");
				} else if (chDesc[u].appliedParagraphStyle.name == "Heading1-Intro-Self") {
					chDesc[u].appliedParagraphStyle = docs[x].paragraphStyles.itemByName("Heading1-Chapter");
				}
				chDesc[u].clearOverrides(OverrideType.ALL);
			}
		} else if (docs[x].name.indexOf("Appendix") != -1) {
			chIcon = "I-chapter_appendix";
		}
		transformChapter(docs[x], chIcon);
	}
}
alert("OS 업데이트 서식 적용 완료합니다.");

function transformChapter(doc, chIcon) {
	var lang = doc.filePath.toString().match(/[^\/\/]+$/g);
	if (lang == "ar-SA" || lang == "he-IL" || lang == "ur-PK") {
		doc.pages[0].textFrames[0].geometricBounds = [30,16,180,184.999999999936]; //텍스트 프레임 크기 조절
	} else {
		doc.pages[0].textFrames[0].geometricBounds = [30,25,180,193.999999999936]; //텍스트 프레임 크기 조절
	}

	var Spds = doc.masterSpreads;
	for (var q=0; q<Spds.length; q++) {
		// 마스터 페이지에 회색 박스 추가
		if (Spds[q].name == "B-Chapter") {
			if (lang == "ar-SA" || lang == "he-IL" || lang == "ur-PK") {
				Spds[q].pages[0].marginPreferences.top = "30mm";
				Spds[q].pages[0].marginPreferences.bottom = "20mm";
				Spds[q].pages[0].marginPreferences.left = "16mm";
				Spds[q].pages[0].marginPreferences.right = "25mm";
			} else {
				Spds[q].pages[0].marginPreferences.top = "30mm";
				Spds[q].pages[0].marginPreferences.bottom = "20mm";
				Spds[q].pages[0].marginPreferences.left = "25mm";
				Spds[q].pages[0].marginPreferences.right = "16mm";
			}
			var backRect = Spds[q].rectangles.add({ 
				geometricBounds : [-4.9999999999999,-5.00000000000011,301.949999999461,214.949999999936],
				fillColor : doc.colors.itemByName("Black"),
				fillTint : 10,
			});
			backRect.sendBackward();
		}
	}
	if (lang == "ar-SA" || lang == "he-IL" || lang == "ur-PK") {
		var chRect = doc.pages[0].rectangles.add({
			geometricBounds : [200,33.5,259.692114427355,85.7111111111111],
			strokeWeight : "0pt",
		});
	} else {
		var chRect = doc.pages[0].rectangles.add({
			geometricBounds : [200.000026914808,119.99996770223,256.822651333279,176.324468994141],
			strokeWeight : "0pt",
		});
	}
	
	var iconFile = new File(doc.filePath + "/images/" + chIcon + ".ai");
	if (!iconFile) {
		alert("챕터 아이콘 파일이 경로에 없습니다. 다시 확인 후에 진행하세요.");
		exit();
	}
	chRect.place(iconFile, false);
	chRect.fit(FitOptions.FRAME_TO_CONTENT);
	chRect.locked = true;
}