function InsertUnicdoefunc() {
	#targetengine "session";
	var w = new Window ("palette","Unicode 입력기");
		group = w.add ('group {orientation: "column"}');
		btn_01 = group.add("button", [0, 0, 200, 25], "RtoL 유니코드 입력");
		btn_02 = group.add("button", [0, 0, 200, 25], "LtoR 유니코드 입력");
	w.show();
	// return CloseWithoutMsg;
	btn_01.onClick = function() {
		insertRtoLUnicode();
	}
	btn_02.onClick = function() {
		insertLtoRUnicode();
	}
}

function functionleftToright() {
	try {
		var result = true;

		if (app.documents.length == 0) {
			alert("인디자인 문서를 열고 실행하세요.");
			exit();
		}
		else if (app.documents.length == 1) {
			directionTransform();
			alert ("바인딩, 스토리, 테이블의 방향 변경 완료")
			return result;
		}
		else if (app.documents.length > 1) {
			var docs = app.documents;
			var myWindow = new Window ('palette', "작업 진행 중 ...");
				myWindow.pbar = myWindow.add ('progressbar', undefined, 0, docs.length);
				myWindow.pbar.preferredSize.width = 300;
			myWindow.show();
			
			for (j=0; j<docs.length; j++) {
				myWindow.pbar.value = j+1;
				var openDocs = app.documents.everyItem().getElements();
				app.activeDocument = openDocs[openDocs.length-1];
				directionTransform();
				// app.activeDocument.save();
			}
			myWindow.close();
			// alert(docs.length + " 개의 문서, 바인딩, 스토리, 테이블의 방향 변경 완료");
			return result;
		}
	} catch(ex) {
		alert(ex);
	}
}

function directionTransform() {
    var doc = app.activeDocument;
    doc.documentPreferences.pageBinding = PageBindingOptions.RIGHT_TO_LEFT;
    for (var s = 0 ; s < doc.stories.length ; s++) {
        var myStory = doc.stories.item(s);
        myStory.storyPreferences.storyDirection = StoryDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
    }
    var a = doc.stories.everyItem().tables.everyItem().getElements();
    if (a.length > 0) {
		for (var i=0; i<a.length; i++){
			a[i].tableDirection = TableDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
		}
		var b = doc.stories.everyItem().tables.everyItem().cells.everyItem().tables.everyItem().getElements();
		if (b.length > 0) {
			for (var j=0; j<b.length; j++){
				b[j].tableDirection = TableDirectionOptions.RIGHT_TO_LEFT_DIRECTION;
			}
		}
	}
}

function allDocusaveClose() {
	var result = true;
	var docs = app.documents;
	
	for (var i = docs.length-1; i >= 0; i--) {
		docs[i].close(SaveOptions.YES);
	}
	// alert("모든 문서 저장 완료");
	return result;
}

function insertRtoLUnicode() {
	if (app.selection[0].constructor.name == "InsertionPoint"){ 
	with (app.selection[0].insertionPoints[0]){ // insertion point 
				contents = String.fromCharCode(0x200E); // dotless i 
	} 
}
}

function insertLtoRUnicode() {
	if (app.selection[0].constructor.name == "InsertionPoint"){ 
	with (app.selection[0].insertionPoints[0]){ // insertion point 
				contents = String.fromCharCode(0x200F); // dotless i 
	} 
}
}