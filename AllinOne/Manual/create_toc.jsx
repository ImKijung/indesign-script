var doc = app.activeDocument;
var myTocPage = doc.spreads[1];
var mTOC = doc.createTOC(doc.tocStyles.item("TOC"), true, null, ["15", "15"])[0];
var myFrame = mTOC.textFrames[0];