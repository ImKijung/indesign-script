function Batch_fit_image_to_frame() {
	/* make by imkijung 20150117  */
	if (app.documents.length == 0) {  
		alert("인디자인 문서를 열고 실행해주세요.");  
	}
	var docs = app.documents;
		for (var i = docs.length-1; i >= 0; i--) {
			fit(docs[i]);
			alert("이미지 프레임에 맞추기 완료");
		}
	function fit(doc) {
		var myGraphics = doc.allGraphics;
		for (d = 0; myGraphics.length > d; d++) {
			if (myGraphics[d].locked == true) {
				continue;
			} else {
				myGraphics[d].parent.fit(FitOptions.proportionally);
				myGraphics[d].parent.fit(FitOptions.centerContent);
			}
		}
	}
}

function Batch_fit_image_size100_frame_to_content() {
	/* make by imkijung 20150117  */
	if (app.documents.length == 0) {  
		alert("인디자인 문서를 열고 실행해주세요.");
	}

	var docs = app.documents;
		for (var i = docs.length-1; i >= 0; i--) {
			fit100(docs[i]);
			alert("이미지 100% 맞추기 완료");
		}
	function fit100(doc) {
		var myGraphics = doc.allGraphics;
		for (d = 0; myGraphics.length > d; d++) {
			if (myGraphics[d].locked == true) {
				continue;
			} else {
				myGraphics[d].horizontalScale = 100;
				myGraphics[d].verticalScale = 100;
				myGraphics[d].parent.fit(FitOptions.FRAME_TO_CONTENT);
			}
		}
	}
}