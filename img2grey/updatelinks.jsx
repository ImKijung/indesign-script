updatelinks();

function updatelinks() {
	var doc = app.activeDocument;
	var imgCount = doc.allGraphics.length;
	var graphics, linkName;
	
	for (var i=0;i<imgCount;i++) {
		graphics = doc.allGraphics[i];
		linkName = graphics.itemLink.name;
		if (linkName.search('.png') > 0 || linkName.search('.jpg') > 0 || linkName.search('.jpeg') > 0) {
			graphics.itemLink.update();
		}
	}
}
