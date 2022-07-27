app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var myFolder = Folder.selectDialog( "현대기아자동차 매뉴얼 전용 XML Export ver.1.2\r\r인디자인 문서가 있는 폴더를 선택하세요." );  
if ( myFolder != null ) {
	var myFiles = [];
	GetSubFolders(myFolder);
	if ( myFiles.length > 0 ) {
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var myBookFileName = myFolder + "/"+ myFolder.name;  
		myBookFile = new File( myBookFileName );
		if ( myBookFile.exists ) {
			if ( app.books.item(myFolder.displayName) == null ) {
				myBook = app.open( myBookFile );
			}
		}
		for ( i=0; i < myFiles.length; i++ ) {
			var myBookFile = app.open(File(myFiles[i]), true);
		}
	}
}
var f = new Folder(myFolder + "/HTML/");
if (!f.exists)
	f.create();

var docs = app.documents;
for (var i=0;i<docs.length;i++){
    var openDocs = app.documents.everyItem().getElements();
    app.activeDocument = openDocs[openDocs.length-1];
	//alert(openDocs[openDocs.length-1].name);
	exportXML();
}
for (var i = docs.length-1; i >= 0; i--) {
	docs[i].close(SaveOptions.NO);
}
// turn on warnings
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
alert ("XML 출력 완료");

function exportXML() {
	var doc = app.activeDocument;

	doc.xmlPreferences.defaultCellTagName = "Cell";
	doc.xmlPreferences.defaultImageTagName = "Image";
	doc.xmlPreferences.defaultStoryTagName = "Story";
	doc.xmlPreferences.defaultTableTagName = "Table";

	doc.deleteUnusedTags();

	//Story에 Tag 적용하기
	var myPages = doc.pages;
	try {
		for(var a = 0; a < myPages.length; a++){
			var myStory = myPages[a].textFrames;
			for (var y=0; y < myStory.length; y++) {
				myStory[y].autoTag();
				}
			}
	} catch(e) {}

	//Lock 설정된 오브젝트 해제하기
	for(var a = 0; a < myPages.length; a++){
		var item = myPages[a].allPageItems;
		for(var b = 0; b < item.length; b++){
			if(item[b].locked){
				item[b].locked = false;
			}
		}
	} 
	//이미지에 Tag 적용하기
	var gStory = doc.stories;
	var myGraphics = doc.allGraphics;
	for (x=0; x < myGraphics.length; x++){
		try {
			myGraphics[x].autoTag();
		} catch(e) {
			// e;
			continue;
		}
	}

	//Enter에 적용된 문자 스타일 해제
	app.findGrepPreferences = app.findChangeGrepOptions = null;
	app.changeGrepPreferences = app.findChangeGrepOptions = null;
	app.findGrepPreferences.properties = {
		findWhat : "\\r",
	}
	app.findChangeGrepOptions.properties = {
		includeFootnotes:false,
		includeHiddenLayers:false,
		includeLockedLayersForFind:false,
		includeLockedStoriesForFind:false,
		includeMasterPages:false,
	}
	app.changeGrepPreferences.appliedCharacterStyle = app.characterStyles.item("$ID/[None]"),
	app.changeGrepPreferences.changeTo = "\\r",
	doc.changeGrep();
	app.findGrepPreferences = app.findChangeGrepOptions = null;
	app.changeGrepPreferences = app.findChangeGrepOptions = null;

	//0번 => 단락 스타일 없슴.
	//1번 => 기본 단락 스타일.
	var myPStyle = doc.allParagraphStyles;
	var myTag = null;
	for (var i = 2 ; i < myPStyle.length ; i++) {
		var myPaStyle = myPStyle[i];// 단락 스타일
		doc.xmlExportMaps.add(myPaStyle, 'para');
	}

	//0번 => 문자 스타일 없슴.
	var myCStyle = doc.allCharacterStyles;
	for (var i = 1 ; i < myCStyle.length ; i++) {
		var myChStyle = myCStyle[i];
		doc.xmlExportMaps.add(myChStyle, myChStyle.name);
	}

	//0번 => 표 스타일 없슴.
	//1번 => 기본 표 스타일.
	var allTableStyles = doc.allTableStyles;
	for (var i = 2 ; i < allTableStyles.length ; i++) {
		var TbStyle = allTableStyles[i];
		doc.xmlExportMaps.add(TbStyle, 'Table');
	}

	//0번 => 셀 스타일 없슴.
	var allCellStyles = doc.allCellStyles;
	for (var i = 1 ; i < allCellStyles.length ; i++) {
		var cellStyle = allCellStyles[i];
		doc.xmlExportMaps.add(cellStyle, 'Cell');
	}

	doc.mapStylesToXMLTags();

	//Hyperlink의 Tag 적용하기
	var myRef = doc.crossReferenceSources.everyItem().getElements();
	for (var i=0; i<myRef.length; i++) {
		app.select(myRef[i].sourceText);
		selection = app.selection[0];
		doc.xmlElements[0].xmlElements.add('xref', selection).xmlAttributes.add('href', String(myRef[i].name));
	}

	//Destination의 Tag 적용하기
	link = doc.hyperlinkTextDestinations;
	for (var i=0 ; i<link.length; i++) {
		var linkName = link[i].name;
		var xxx = link[i].destinationText;
		xxx.select();
		selection = app.selection[0];
		doc.xmlElements[0].xmlElements.add('anchor', selection).xmlAttributes.add('x', String(linkName));
	}

	// //Element 별 @style 속성 적용하기
	var rootXML = doc.xmlElements[0];

	//para @style 적용
	var paraElement = rootXML.evaluateXPathExpression (".//para");
	for (var i=0; i<paraElement.length; i++) {
		app.select(paraElement[i]);
		sel = app.selection[0];
		paraStyle = sel.paragraphs[0].appliedParagraphStyle;
		paraElement[i].xmlAttributes.add('style', String(paraStyle.name));
	}

	// //span @style 적용
	// var spanElement = rootXML.evaluateXPathExpression (".//span");
	// for (var i=0; i<spanElement.length; i++) {
	// 	app.select(spanElement[i]);
	// 	sel = app.selection[0];
	// 	myCStyle = sel.insertionPoints[0].appliedCharacterStyle;
	// 	spanElement[i].xmlAttributes.add('style', String(myCStyle.name));
	// }

	//Table @style 적용
	var tbElement = rootXML.evaluateXPathExpression (".//Table");
	for (var i=0; i<tbElement.length; i++) {
		app.select(tbElement[i]);
		sel = app.selection[0];
		TbStyle = sel.tables[0].appliedTableStyle;
		tbElement[i].xmlAttributes.add('style', String(TbStyle.name));
	}

	//Cell @style 적용
	var cellElement = rootXML.evaluateXPathExpression (".//Cell");
	for (var i=0; i<cellElement.length; i++) {
		app.select(cellElement[i]);
		sel = app.selection[0];
		ClStyle = sel.cells[0].appliedCellStyle;
		cellElement[i].xmlAttributes.add('style', String(ClStyle.name.split("없음").join("None")));
	}

	//Table과 Cell의 상관관계 정리
	var tables = doc.xmlElements[0].evaluateXPathExpression(".//Table");
	var tables_count = tables.length;
	if (tables_count > 0) {
		for (var i=0 ; i<tables_count ; i++) {
			var table = tables[i].tables[0];
			tables[i].xmlAttributes.add('HeaderRowCount', table.headerRowCount + '');
			var rows = table.rows;
			var rows_count = rows.length;
			for (var j=0 ; j<rows_count ; j++)
				rows[j].cells[0].associatedXMLElement.xmlAttributes.add('newrow', 'newrow');
			var columns = table.columns;
			var columns_count = columns.length;
			var widths = [];
			for (var j=0 ; j<columns_count ; j++)
				widths = widths.concat((columns[j].width/table.width).toFixed(2) + '*');
				tables[i].xmlAttributes.add('colspecs', widths.join(':'));
		}
		var cells = doc.xmlElements[0].evaluateXPathExpression(".//Cell");
		var cells_count = cells.length;
		for (var i=0 ; i<cells_count ; i++) {
			var cell = cells[i].cells[0];
			if (cell.columnSpan > 1) {
				var namest = cell.parentColumn.index + 1;
				var nameend = namest + cell.columnSpan - 1;
				cells[i].xmlAttributes.add('namest', 'col' + namest);
				cells[i].xmlAttributes.add('nameend', 'col' + nameend);
			}
		}
	}

	//XML 출력
	var myName = doc.name.split(".indd").join(".xml");
	app.scriptPreferences.userInteractionLevel=UserInteractionLevels.neverInteract;
	var myInxfile = new File(doc.filePath.absoluteURI + "/HTML/" + doc.name.split(".indd")[0]+".xml");
	doc.exportFile(ExportFormat.XML, myInxfile);
	//doc.close(SaveOptions.NO)
	//alert("XML Export 완료")
}
//=================================== FUNCTIONS =========================================
function GetSubFolders(theFolder) {
	var myFileList = theFolder.getFiles();
	for (var i = 0; i < myFileList.length; i++) {
		var myFile = myFileList[i];
		if (myFile instanceof Folder){
			GetSubFolders(myFile);
		}
		else if (myFile instanceof File && myFile.name.match(/\.indd$/i)) {
			myFiles.push(myFile);
		}
	}
}
//========================================================================================