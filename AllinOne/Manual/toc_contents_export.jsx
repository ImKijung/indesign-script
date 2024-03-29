function toc_contents_export() {
	if (app.documents.length == 0) {  
			alert("인디자인 문서를 열고 실행해주세요.");  
			exit();  
	}
		
	var myDiag = new Window( 'dialog',"단락스타일 컨텐츠 추츨하기",undefined,{closeButton:true} );
		myDiag.alignChildren = "right";
		myDiag.minimumSize.width = 300;
	var main = myDiag.add("group");
		main.add("statictext",undefined,"단락스타일1:");
	var group_1 = main.add("group{alignChildren:'left',orientation:'stack'}");
	var list_1 = group_1.add('dropdownlist'); 
	var paraStyles_1 = app.activeDocument.allParagraphStyles;
	var listItems_1 = [];
	
	for( var l = 0; l < paraStyles_1.length; l++ ) {
		var listItems_1 = list_1.add('item', paraStyles_1[l].name);
	}
	main.add("statictext",undefined,"단락스타일2:");
	
	var group_2 = main.add("group{alignChildren:'left',orientation:'stack'}");
	var list_2 = group_2.add('dropdownlist'); 
	var paraStyles_2 = app.activeDocument.allParagraphStyles;
	var listItems_2 = [];
	
	for( var m = 0; m < paraStyles_2.length; m++ ) {
		var listItem_2 = list_2.add('item', paraStyles_2[m].name);
	}

	var export_button = myDiag.add('button{text:"Export"}')

	export_button.onClick = function() {
		//alert(list.selection.text)
		var psName_1 = list_1.selection.text;
		var psName_2 = list_2.selection.text;
		var curDoc = app.activeDocument;  
		var docPath = curDoc.filePath;   
	
		// create a file  
		var fileToWrite = new File(docPath +"/" + curDoc.name.replace(/\.indd$/,"")  + "_paragraph" + ".txt");  
		var ok = fileToWrite.open("w");  
			if (!ok) {  
				alert("Error: " + fileToWrite.error);  
				exit();  
		}
		var csv = File (fileToWrite);
		csv.encoding = "UTF-8"; //Encoding for Windows/Excel
		csv.open ("w");
		// collect the content and write it to the file  
		var allPages= app.activeDocument.pages;
		for (i = 0; i < allPages.length; i++) {
		var curPage_number = allPages[i].name;
		var allPageTextframes= allPages[i].textFrames;
		for (j = 0; j < allPageTextframes.length; j++) {
			var allPageTextframeParagraphs= allPageTextframes[j].paragraphs;
			for (k = 0; k < allPageTextframeParagraphs.length; k++) {
				var text = allPageTextframeParagraphs[k].contents;
				var paragraphStyleName = allPageTextframeParagraphs[k].appliedParagraphStyle.name;
				switch(allPageTextframeParagraphs[k].appliedParagraphStyle.name) {
					case psName_1:
						toc_contents = allPageTextframeParagraphs[k].contents;
						toc_con = toc_contents.replace("\x{FEFF}", "");
					case psName_2:
						toc_contents = allPageTextframeParagraphs[k].contents;
						toc_con = toc_contents.replace("\x{FEFF}", "");
					var zeile = paragraphStyleName + "\t" + curPage_number + "\t" + toc_con;
					csv.writeln (zeile);
				}
			}
		}
	}
	fileToWrite.close();
	fileToWrite.execute();
	myDiag.close();
	}
	myDiag.show();
}