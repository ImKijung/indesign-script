var curDoc = app.activeDocument;

var curCont = curDoc.pages[0].textFrames[0].paragraphs;

for (var i=0;i<curCont.length;i++) {
	if (curCont[i].contents.match("전자파 흡수율")) {
		var count = i;
		break;
	}
}

var korTf = curDoc.textFrames.add({
	geometricBounds : [25.0000000018403,15.9999983344793,36,193.999998347582],
});

var Conts = curDoc.pages[0].textFrames[1].paragraphs;

for (var j=count-1; j>=0; j--) {
	Conts[j].select();
	app.cut();
	korTf.insertionPoints[0].select();
	app.paste();
}

korTf.textFramePreferences.insetSpacing = [ 4, 4, 4, 4 ];
korTf.strokeWeight = 0.5;
korTf.fit(FitOptions.FRAME_TO_CONTENT);