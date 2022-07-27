function all_save_close_SERIAL2() {
	if(app.documents.length == 0){
		alert("인디자인 문서를 여세요!");
		exit();
	} else {
		mainASCS ();
	}

	//===================================== Save|1512460  ======================================
	function mainASCS() {
		var docs = app.documents;
		var myMenuAction = app.menuActions.itemByID (1512460);
		for (var i = docs.length-1; i >= 0; i--) {
			try {
				var myActionCheck = myMenuAction.enabled
			} catch (e) {
				alert ("");
			}
			
			if (myActionCheck == false) {
				alert ("현재 문서는 수정된 사항이 없기 때문에 문서를 저장하지 않고 종료합니다.")
				app.activeDocument.close(SaveOptions.NO);
			} 
			if (myActionCheck == true) {
				myMenuAction.invoke();
				app.activeDocument.close(SaveOptions.YES);
			}
		}
	}
}