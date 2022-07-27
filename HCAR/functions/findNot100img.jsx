function CheckImageSize(myDoc) {
	try {
		var allImages = myDoc.allGraphics;
		var imageCount = GetImageCountWithoutMaster(myDoc);
		var report = [];
		var count = 0;
		for (var i = 0; i < imageCount; i++) {
			var image = allImages[i];
			var page = '알수 없는 페이지';
			if (image.parentPage != null)
				page = image.parentPage.name + 'page';
			if (image.absoluteVerticalScale != 100 || image.absoluteHorizontalScale != 100) {
				report.push(myDoc.name + ' : ' + page + ' - 세로:'
					+ image.absoluteHorizontalScale + ' | 가로:' + image.absoluteVerticalScale + ' : ' + image.itemLink.name);
				count ++;
			}
		}
		if (count > 0) {
			var docPath = myDoc.filePath;
			var Report = new File(docPath +"/" + myDoc.name.replace(".indd","") + "_images-Report.txt");
			Report.open("w");
			Report.encoding = "utf-8";
			Report.writeln(String(report).replace(/\,/g, "\r"));
			Report.close();
		}
	} catch (ex) {
		alert('CheckImageSize 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
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