var doc = app.activeDocument;
var docPath = doc.filePath;
// var file = new File(docPath + "/" + doc.name.split(".indd").join(".txt"));
// file.open("w");

// var Pages= dpc.pages;
// for (i = 0; i < Pages.length; i++) {
// 	var tf= Pages[i].textFrames;
// 	for (j = 0; j < tf.length; j++) {
// 		var myPara = tf[j].paragraphs;
// 		for (k = 0; k < myPara.length; k++) {
// 			var text = myPara[k].contents;
// 			var paragraphStyleName = myPara[k].appliedParagraphStyle.name;
// 			//$.writeln (paragraphStyleName);
// 			//$.writeln (text);
// 			var paraCont = paragraphStyleName + "\t" + text;
			
// 			for (l=0;l<)
// 			// file.writeln (zeile);
// 			$.writeln(paraCont)
// 		}
// 	}
// }

var myPage = doc.pages;
for (var i=0;i<myPage.length;i++) {
	$.writeln(myPage[i].name);
	var myPara = myPage.everyItem().getElements();
	for (var j=0;j<myPara.length;j++) {
		$.writeln(myPara[j].name);
	}
}