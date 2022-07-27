var doc = app.activeDocument;
	var updateLayer = doc.layers.itemByName("업데이트");
	if (updateLayer.isValid) {
		alert("ㅇㅇ");
		updateLayer.visible = false;
	} else 
		alert("GG");