if(app.documents.length != 0){
    var myFolder = Folder.selectDialog ('블루투스 아이콘이 있는 이미지 폴더를 선택하세요.');
    if(myFolder != null){
        // reset the Find/Change dialog
        app.findGrepPreferences = app.changeGrepPreferences = null;
        // formulate a grep search string
        app.findGrepPreferences.findWhat = 'º';
        // find all occurrence, last one first
        f = app.activeDocument.findGrep (true);
        for (i = 0; i < f.length; i++){
            // construct file name
            name = f[i].contents.replace (/@/g, '');
            
            // place the image
            var tempRect = f[i].insertionPoints[0].rectangles.add( );
            tempRect.place (File (myFolder.fsName + '/' + 'ico_BT.ai'));
            tempRect.fit(FitOptions.FRAME_TO_CONTENT);
            tempRect.topLeftCornerOption = CornerOptions.NONE;
            tempRect.topRightCornerOption = CornerOptions.NONE;
            tempRect.bottomLeftCornerOption = CornerOptions.NONE;
            tempRect.bottomRightCornerOption = CornerOptions.NONE;
            tempRect.strokeWeight = "0pt";
            tempRect.strokeColor = "None";
            
        }
    // delete all @??@ codes
    app.activeDocument.changeGrep();
    }
}
else{
   alert('Please open a document and try again.');
}

(function(){  
	if (app.documents.length > 0) {  
		var anchArray = [  
			// edit names; add new pairs as needed  
			{'name' : "ico_BT.ai", 'char' : "§"}, 
			// my test case  
			]  
		replaceAnchors(app.documents[0], anchArray);  
	}  


	function replaceAnchors(aDoc, anchArray) {  
		var links = aDoc.links.everyItem().getElements();  
			for (var j = links.length - 1; j >= 0; j--) {  
		var anchor = getAnchor(links[j]);  
			if (anchor == null) continue; // link not anchored  
		var aChar = getCharacter(links[j].name, anchArray);  
			if (aChar != null) {  
				//anchor.contents = aChar;  
				anchor.appliedCharacterStyle = "C_Image"  
			}  
		}  
	}  


	function getCharacter(linkName, anchArray) {  
		for (var j = anchArray.length - 1; j >= 0; j--) {  
			if (linkName == anchArray[j].name) {  
				return anchArray[j]['char'];  
			}  
		}  
		return null;  
	}  


	function getAnchor(link) {  
		var theChar = link.parent.parent.parent;  
			if (theChar instanceof Character) {  
				return theChar;  
			}  
			return null; // not anchored  
		}  
	}())