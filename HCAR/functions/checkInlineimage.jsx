function checkInlineImages() {
	var myWin = new Window ("palette", "인라인 이미지 크기 검사");
		myWin.current = myWin.add ('statictext {text: "검사 개수: ", characters: 30, justify: "center"}');
		myWin.button = myWin.add("button", [0, 0, 180, 25], "시작");

	var state = 0;
	myWin.current.text = "검사 개수: " + state;

	myWin.show();

	myWin.button.onClick = function () {
		var myDoc = app.activeDocument;
		var allImages = myDoc.allGraphics;
		var imageCount = GetImageCountWithoutMaster(myDoc);
		if (imageCount == 0) {
			alert("이 문서에는 이미지가 없습니다.");
			exit();
		}
		if (state == imageCount) {
			state = 0;
		}
		for (var i=state; i<imageCount; i++) {
			var image = allImages[i];
			state = i + 1;
			myWin.current.text = " 검사 개수: " + state + " / " + imageCount;
			if (image.itemLink.name.indexOf("ico_") != -1) {
				if (image.absoluteVerticalScale != 100 || image.absoluteHorizontalScale != 100) {
					image.select();
					app.activeWindow.activePage = image.parent.parent.parentTextFrames[0].parentPage;
					app.activeDocument.layoutWindows[0].zoomPercentage = 300;
					// break;
					// w.close();
					fitimg();
				}
			}
			if (state == imageCount) {
				state = 0;
				myWin.current.text = " 검사 개수: " + state + " / " + imageCount;
				alert("완료합니다.");
				exit();
			}
		}
	}
}

function fitimg() {
	var cfwin = new Window("dialog");
	cfwin.add ('statictext {text: "이미지를 100%로 수정하시겠습니까?", justify: "center"}');
	cfwin.add("button", undefined, "OK",{name:"ok"});
	cfwin.add("button", undefined, "Cancel",{name:"cancel"});
	
	var result = cfwin.show();

	if (result == 1) {
		var mySel = app.selection[0];
		mySel.horizontalScale = 100;
		mySel.verticalScale = 100;
		mySel.parent.fit(FitOptions.FRAME_TO_CONTENT);
	}
}

function GetImageCountWithoutMaster(myDoc) {
	try {
		var myMaster = myDoc.masterSpreads;
		var myMasterNumber = myMaster.count();
		var masterImages = 0;
		var imageNumber;
		for (var i = 0 ; i < myMasterNumber ; i++) {
			imageNumber = myMaster[i].allGraphics;
			masterImages += imageNumber.length;
		}

		return myDoc.allGraphics.length - masterImages;
	} catch (ex) {
		alert(ex);
	}
}