function ChagneFont() {
	if (app.documents.length < 1) {
		alert("인디자인 문서를 여세요!");
		exit();
	} else {
		mainChangeFont()();
	}
}

function mainChangeFont() {
	var doc = app.activeDocument;

	var myDialog= new Window("dialog","Font replacer");
	var grp1=myDialog.add("group");
	var panel1=grp1.add("panel",undefined,"Find font:");
	var panel2=grp1.add("panel",undefined,"Change font:");
	var findFonts=doc.fonts.everyItem().name.join("\r").replace(/\t/g,":").split("\r");
	var chFonts=app.fonts.everyItem().name.join("\r").replace(/\t/g,":").split("\r");
	var myFindDrop=panel1.add("dropdownlist",undefined,findFonts);
	var myChDrop=panel2.add("dropdownlist",undefined,chFonts);

	myDialog.add("button",undefined,"Ok",{name:"ok"});
	myDialog.add("button",undefined,"Cancel",{name:"cancel"});
	myFindDrop.selection=myFindDrop.items[0];
	myChDrop.selection=myChDrop.items[0];
		
	if(myDialog.show() != 1){
		exit()
		}else {
			var pstyles = doc.allParagraphStyles;
			var cstyles = doc.allCharacterStyles;
			var ffont = myFindDrop.selection.text.replace(/:/g,"\t");
			var tfont = myChDrop.selection.text.replace(/:/g,"\t");
		for (var a = 1; a < pstyles.length; a++) {
			switch (pstyles[a].appliedFont.name) {
				case ffont: pstyles[a].appliedFont = tfont; break;
			}
		};

		for (var i = 1; i < cstyles.length; i++) {
			if ((cstyles[i].appliedFont+'\t'+cstyles[i].fontStyle==ffont)||(cstyles[i].appliedFont.name==ffont)) {
				cstyles[i].appliedFont = tfont;
				}
			};
		}
}