function functionCaptionLabel() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '도안 번호 작업 중: ';
    return function_core(CaptionLabel, 'multi', true, false);
}

function CaptionLabel(myDoc) {
	try {
	  	var pStyles = myDoc.allParagraphStyles;
		for (var i=1; i<pStyles.length; i++) {
			var ps = pStyles[i];
			switch (ps.name) {
				case "도안번호" : ps.justification = 1818584692; break; //left
			}
		}
		var myLinks = myDoc.links;
		var myLink, myLinkPos, myLinkName, myLinkPath, myLinkType, myLinkStat, myLinkInfoStr,
			myLinkParent, myLinkSpread, myLinkStatus, myLinkInfoFrame;
		
		for (var j=0; j<myLinks.length; j++) {
			myLink = myLinks[j];
			var myLink_name = myLink.name;
			if (myLink_name.indexOf("D-") != -1) {
				if (myLink.parent.parent.itemLayer.visible == true) {
					myLinkPos = myLink.parent.parent.geometricBounds;
					// myLinkName = "";
					// myLinkPath = "";
					// myLinkType = "";
					// myLinkStat = "";
					// myLinkInfoStr = "";
				
					myLinkParent = myLink.parent;
					do {
						if (myLinkParent.constructor.name != "Character") {
							myLinkParent = myLinkParent.parent;
						}
						else {
							if (app.version.split(".")[0] >= 4) {
								myLinkParent = myLinkParent.parentTextFrames[0];
							}
							else {
								myLinkParent = myLinkParent.parentTextFrame;
							}					
						}
					} while ((myLinkParent.constructor.name != "Spread")&&(myLinkParent.constructor.name != "MasterSpread"))
					if(myLinkParent.constructor.name == "Spread"){
						myLinkSpread = myDoc.spreads[myLinkParent.index];
					}
					else{
						myLinkSpread = myDoc.masterSpreads[myLinkParent.index];
					}
				
					if(myLink_name.checkedState){
						if((myPathOpt.checkedState == false) && (myTypeOpt.checkedState == false) && (myStatusOpt.checkedState == false)){
							myLinkInfoStr = myLinkInfoStr + myLink.name
						}
						else{
							myLinkInfoStr = myLinkInfoStr + "Name:\t" + myLink.name
						}
					}
					myDoc.activeLayer = myDoc.layers.itemByName("도안번호");	
					myLinkInfoFrame = myLinkSpread.textFrames.add();
					myLinkInfoFrame.geometricBounds = [myLinkPos[2],myLinkPos[1],myLinkPos[2]+10,myLinkPos[1]+50];
					// y 좌표값 myLinkPos[0] , x 좌표값 myLinkPos[1]
					myLinkInfoFrame.textFramePreferences.insetSpacing = ["0mm","0mm","0mm","0mm"];
					myLinkInfoFrame.textFramePreferences.ignoreWrap = true;
					// myLinkInfoFrame.fillColor = myInfoTitle;
					myLinkInfoFrame.fillTint = 25;
					// myLinkInfoFrame.strokeColor = myInfoTitle;
					// myLinkInfoFrame.strokeWeight = 1;
					// if(myLabelOpt.checkedState){
					// 	myLinkInfoFrame.nonprinting = false;
					// }
					// else{
					// 	myLinkInfoFrame.nonprinting = true;
					// }

					myLinkInfoFrame.contents = myLink_name;
					for(p = 0;p < myLinkInfoFrame.parentStory.paragraphs.length;p++){
						myLinkInfoFrame.parentStory.paragraphs[p].appliedParagraphStyle = "도안번호";
						// myLinkInfoFrame.parentStory.paragraphs[p].words[0].appliedCharacterStyle = myInfoTitle;
					}
					myLinkInfoFrame.fit(FitOptions.frameToContent); 
				}
			}
		}
    	return true;
	}
	catch(ex){
		errorMsgs.push('CaptionLabel 함수 오류(' + myDoc.name + '):Line:' + ex.line + '::' + ex);
        hasError = true;
        throw ex;
	}
}