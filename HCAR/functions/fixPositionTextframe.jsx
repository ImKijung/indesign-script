function fixTextFrames(myDoc) {
	var pageWidth = myDoc.documentPreferences.pageWidth;
	var myPages = myDoc.pages;
	var marginlf, allTf, yPosition, moveVal, pageNumber;
	for (var i=0; i<myPages.length; i++) {
		pageNumber = myPages[i].name.split("-")[1];
		// 안쪽 마진을 구한다.
		marginlf = myPages[i].marginPreferences.left;
		
		// 텍스트 프레임의 Y 값을 구한다.
		allTf = myPages[i].textFrames;
		for (var j=0; j<allTf.length; j++) {
			if (pageNumber % 2 == 0) {
				// 페이지가 짝수 일때는 y + 8
				yPosition = allTf[j].geometricBounds[1];
				yPosition = yPosition + 8;
				if (yPosition != marginlf) {
					moveVal = marginlf - yPosition
					allTf[j].move(undefined, [moveVal, 0])
				}
			} else {
				// 페이지가 홀수 일때는 y 좌표
				yPosition = allTf[j].geometricBounds[1];
				if (yPosition != marginlf) {
					if (i == 0) {
						moveVal = marginlf - yPosition
						allTf[j].move(undefined, [moveVal, 0]);
					} else {
						moveVal = marginlf - yPosition + pageWidth;
						// $.writeln(pageNumber + " - " + moveVal);
						allTf[j].move(undefined, [moveVal, 0]);
					}
				}
			}
		}	
	}
}