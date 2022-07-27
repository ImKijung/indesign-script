var doc = app.activeDocument;
var zoomSet = 400;
var Empty2 = [ "Step-Cmd-OL1", "Step-Cmd-OL2", "Step-Description_2", "Step-UL1_2", "Step-UL1_2-Note", "UL1-Caution-Warning", "UL1_1-Note", "Empty_2" ];
var Empty3 = "Step-UL1_2";

var pages = doc.pages;
var tf, paras, pStyle, prevStyle;

for (var i=0; i<pages.length; i++) {
	var tf = pages[i].textFrames;
	for (var j=0; j<tf.length; j++) {
		var paras = tf[j].paragraphs;
		for (var k=0; k<paras.length; k++) {
			curPara = paras[k];
			pStyle = paras[k].appliedParagraphStyle.name;
			if (k > 0) {
				prevStyle = paras[k-1].appliedParagraphStyle.name;
			}
			if (pStyle == "Empty_2") {
				if (chkMatchStyle2(prevStyle)) {
					curPara.select();
					app.activeWindow.zoomPercentage = zoomSet;
					alert(prevStyle + "+" + pStyle + " 스타일이 잘못 적용되어 있습니다.");
					exit();
				}
			} else if (pStyle == "Empty_3") {
				if (chkMatchStyle3(prevStyle)) {
					curPara.select();
					app.activeWindow.zoomPercentage = zoomSet;
					alert(prevStyle + "+" + pStyle + " 스타일이 잘못 적용되어 있습니다.");
					exit();
				}
			}
		}
	}
}
$.writeln("Clean~~");

function chkMatchStyle2(pStyle) {
	var count = 0;
	for (var i=0;i<Empty2.length;i++) {
		if (Empty2[i] == pStyle ) {
			count ++;
		}
	}
	if (count == 0) {
		return true;
	} else return false
}

function chkMatchStyle3(pStyle) {
	var count = 0;
	for (var i=0;i<Empty3.length;i++) {
		if (Empty3[i] == pStyle ) {
			count ++;
		}
	}
	if (count == 0) {
		return true;
	} else return false
}
