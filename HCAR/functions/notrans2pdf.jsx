function _notrans2pdf(doc, exportPath) {
	// var doc = app.activeDocument;
	var noTrs = [ "C_Notrans_NoBold", "C_Notrans" ];
	var cStyles = doc.allCharacterStyles;

	//색상 만들기
	var myColorToCheckProperties = {
		name : "check_notrans",
		colorValue : [0,100,0,0],
		space : ColorSpace.CMYK
	};

	// 색상이 있는지 없는지 확인하기
	var myColorToCheck = doc.colors.itemByName(myColorToCheckProperties.name);

	if (!myColorToCheck.isValid) {
		doc.colors.add (myColorToCheckProperties);
	}

	// notrans 문자 스타일에 추가한 색상 적용하기
	for (var i=1; i<cStyles.length; i++) {
		var cs = cStyles[i];
		for (var j=0; j<noTrs.length; j++) {
			if (cs.name == noTrs[j]) {
				cs.fillColor = "check_notrans";
			}
		}
	}

	// PDF를 내보낸다. 사전설정, AST_CAR_Manual
	var pdfPreset = app.pdfExportPresets.item("AST_CAR_Manual");

	var outPDF = new File(exportPath + "/" + doc.name.replace('.indd', '.pdf'));
	doc.exportFile(ExportFormat.pdfType, outPDF, false, pdfPreset);
	ChangeNone(doc, cStyles, noTrs);
}

function ChangeNone(doc, cStyles, noTrs) {
	for (var i=1; i<cStyles.length; i++) {
		var cs = cStyles[i];
		for (var j=0; j<noTrs.length; j++) {
			if (cs.name == noTrs[j]) {
				cs.fillColor = NothingEnum.NOTHING;
			}
		}
	}
}