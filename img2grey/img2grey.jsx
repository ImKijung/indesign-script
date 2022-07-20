grey2img();

function grey2img() {
	var doc = app.activeDocument;
	var docPath = doc.filePath;
	var imgCount = doc.allGraphics.length;
	var graphics, linkName, convertName
	
	for (var i=0;i<imgCount;i++) {
		graphics = doc.allGraphics[i];
		linkName = graphics.itemLink.name;
		linkPath = graphics.itemLink.filePath;
		if (linkName.search('.png') > 0 || linkName.search('.PNG') > 0) {
			app.pngExportPreferences.pngColorSpace = PNGColorSpaceEnum.GRAY
			app.pngExportPreferences.pngQuality = PNGQualityEnum.HIGH;
			app.pngExportPreferences.transparentBackground = true;
			convertName = new File(linkPath);
			graphics.exportFile(ExportFormat.PNG_FORMAT, convertName);
		} else if (linkName.search('.jpg') > 0 || linkName.search('.jpeg') > 0 || linkName.search('.JPG') > 0 || linkName.search('.JPEG') > 0) {
			app.jpegExportPreferences.jpegColorSpace = JpegColorSpaceEnum.GRAY
			app.jpegExportPreferences.jpegQuality = JPEGOptionsQuality.HIGH
			convertName = new File(linkPath);
			graphics.exportFile(ExportFormat.JPG, convertName);
		}
	}
	alert("완료합니다. updatelinks 스크립트를 실행하세요.");
}
