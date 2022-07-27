function Hyperlink_apply_underline() {
	if (app.documents.length == 0) {  
		alert("인디자인 문서를 열고 실행해주세요.");  
		exit();  
	}
	
	var docs = app.documents;  
	
	for (var i = 0; i < docs.length; i++) {  
		findapplyHyp(docs[i]);  
		}  
		alert("완료합니다.");  
	
	function findapplyHyp(doc) {  
		var applyStyle = doc.characterStyles.item("C_Font-Underline");  
		var myHyp = doc.hyperlinks.everyItem().getElements();  
		var len = myHyp.length;  

		while (len-->0) {   
			myHyp[len].source.sourceText.appliedCharacterStyle = applyStyle;  
		}
	}
}