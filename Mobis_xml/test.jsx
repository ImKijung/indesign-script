var doc = app.activeDocument;

for (var d = doc.links.length-1; d >= 0; d--) {
	link = doc.links[d];
	if (link.name.indexOf(".PNG") != -1) {
		newName = link.name.replace(".PNG", ".png");
		var curLinkFile = new File(link.filePath);
		curLinkFile.rename(newName);
		link.relink(curLinkFile);
	}
	if (link.name.indexOf(".jpg") != -1 || link.name.indexOf(".JPG") != -1) {
		var pngName = link.filePath.replace(/\.(jpg|JPG)/gi, ".png");
		var pngFile = new File(pngName);
		$.writeln(pngFile);
		link.parent.exportFile(ExportFormat.PNG_FORMAT, pngName);
		link.relink(pngFile);
	}
}