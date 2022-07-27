app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var myFolder = Folder.selectDialog( "현대기아자동차 매뉴얼 전용 번역 XML Export [dev]\r\r인디자인 문서가 있는 폴더를 선택하세요." );  
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
} else if ( myFolder == null) {
	exit();
}

var docs = app.documents;
var f = new Folder(myFolder + "/Translate/");
if (!f.exists)
	f.create();

for (var i=0;i<docs.length;i++){
    var openDocs = app.documents.everyItem().getElements();
    var myDoc = app.activeDocument = openDocs[openDocs.length-1];
	//alert(openDocs[openDocs.length-1].name);
	exportXML(myDoc);
}
for (var i = docs.length-1; i >= 0; i--) {
	docs[i].close(SaveOptions.NO);
}
// turn on warnings
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
alert ("XML 출력 완료");

function exportXML(myDoc) {
	// var doc = app.activeDocument;
	var updateLayer = myDoc.layers.itemByName("업데이트");
	if (updateLayer.isValid) {
		updateLayer.visible = false;
	} else

	myDoc.xmlPreferences.defaultCellTagName = "Cell";
	myDoc.xmlPreferences.defaultImageTagName = "Image";
	myDoc.xmlPreferences.defaultStoryTagName = "Story";
	myDoc.xmlPreferences.defaultTableTagName = "Table";

	myDoc.deleteUnusedTags();

	//Story에 Tag 적용하기
	var myPages = myDoc.pages;
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
	var gStory = myDoc.stories;
	var myGraphics = myDoc.allGraphics;
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
	myDoc.changeGrep();
	app.findGrepPreferences = app.findChangeGrepOptions = null;
	app.changeGrepPreferences = app.findChangeGrepOptions = null;

	//0번 => 단락 스타일 없슴.
	//1번 => 기본 단락 스타일.
	var myPStyle = myDoc.allParagraphStyles;
	var myTag = null;
	for (var i = 2 ; i < myPStyle.length ; i++) {
		var myPaStyle = myPStyle[i];// 단락 스타일
		myDoc.xmlExportMaps.add(myPaStyle, 'para');
	}

	//0번 => 문자 스타일 없슴.
	var myCStyle = myDoc.allCharacterStyles;
	for (var i = 1 ; i < myCStyle.length ; i++) {
		var myChStyle = myCStyle[i];
		myDoc.xmlExportMaps.add(myChStyle, 'span'); //변경
	}

	//0번 => 표 스타일 없슴.
	//1번 => 기본 표 스타일.
	var allTableStyles = myDoc.allTableStyles;
	for (var i = 2 ; i < allTableStyles.length ; i++) {
		var TbStyle = allTableStyles[i];
		myDoc.xmlExportMaps.add(TbStyle, 'Table');
	}

	//0번 => 셀 스타일 없슴.
	var allCellStyles = myDoc.allCellStyles;
	for (var i = 1 ; i < allCellStyles.length ; i++) {
		var cellStyle = allCellStyles[i];
		myDoc.xmlExportMaps.add(cellStyle, 'Cell');
	}

	myDoc.mapStylesToXMLTags();

	//Hyperlink의 Tag 적용하기
	var myRef = myDoc.crossReferenceSources.everyItem().getElements();
	for (var i=0; i<myRef.length; i++) {
		app.select(myRef[i].sourceText);
		selection = app.selection[0];
		myDoc.xmlElements[0].xmlElements.add('xref', selection).xmlAttributes.add('href', String(myRef[i].name));
	}

	//Destination의 Tag 적용하기
	link = myDoc.hyperlinkTextDestinations;
	for (var i=0 ; i<link.length; i++) {
		var linkName = link[i].name;
		var xxx = link[i].destinationText;
		xxx.select();
		selection = app.selection[0];
		myDoc.xmlElements[0].xmlElements.add('anchor', selection).xmlAttributes.add('x', String(linkName));
	}

	// //Element 별 @style 속성 적용하기
	var rootXML = myDoc.xmlElements[0];
	
	//문장, 문자 스타일 속성 추가
	var paras = myDoc.xmlElements[0].evaluateXPathExpression(".//para");
	var paras_count = paras.length;
	if (paras_count > 0) {
		for (var i=0;i<paras_count;i++) {
			var para = paras[i].paragraphs[0];
			paras[i].xmlAttributes.add('style', para.appliedParagraphStyle.name + '');
		}
	}
	
	var spans = myDoc.xmlElements[0].evaluateXPathExpression(".//span");
	var spans_count = spans.length;
	if (spans_count > 0) {
		for (var j=0;j<spans_count;j++) {
			var span = spans[j].textStyleRanges[0];
			spans[j].xmlAttributes.add('style', span.appliedCharacterStyle.name + '');
		}
	}

	//Table과 Cell에 속성 추가
	var tables = myDoc.xmlElements[0].evaluateXPathExpression(".//Table");
	var tables_count = tables.length;
	if (tables_count > 0) {
		for (var i=0 ; i<tables_count ; i++) {
			var table = tables[i].tables[0];
			var tbDirection = "";
			tables[i].xmlAttributes.add('HeaderRowCount', table.headerRowCount + '');
			//추가 속성 by iM - 20.07.30
			tables[i].xmlAttributes.add('FooterRowCount', table.footerRowCount + '');
			tables[i].xmlAttributes.add('BodyRowCount', table.bodyRowCount + '');
			tables[i].xmlAttributes.add('ColumnCount', table.columnCount + '');
			tables[i].xmlAttributes.add('Width', table.width * 2.835 + '');
			if (table.tableDirection == "1278366308") {
				tbDirection = "LeftToRightDirection";
			} else
				tbDirection = "RightToLeftDirection";
			tables[i].xmlAttributes.add('TableDirection', tbDirection + '');
			tables[i].xmlAttributes.add('AppliedTableStyle', table.appliedTableStyle.name + '');
			//추가 속성 by iM - 20.07.30
			var rows = table.rows;
			var rows_count = rows.length;
			var rowHeight = [];
			for (var j=0 ; j<rows_count ; j++) {
				rows[j].cells[0].associatedXMLElement.xmlAttributes.add('newrow', 'newrow');
				//추가 속성 by iM - 20.08.05
				rowHeight = rowHeight.concat(rows[j].cells[0].height * 2.835 + "");
			}
			tables[i].xmlAttributes.add('RowHeight', rowHeight.join('|'));
			//추가 속성 by iM - 20.08.05
			var columns = table.columns;
			var columns_count = columns.length;
			var widths = [];
			for (var j=0 ; j<columns_count ; j++)
				widths = widths.concat((columns[j].width/table.width).toFixed(2) + '*');
				tables[i].xmlAttributes.add('colspecs', widths.join(':'));
		}
		var cells = myDoc.xmlElements[0].evaluateXPathExpression(".//Cell");
		var cells_count = cells.length;
		for (var i=0 ; i<cells_count ; i++) {
			var cell = cells[i].cells[0];
			var cellfBaseline = "";
			var cellVerticajust = "";
			if (cell.columnSpan > 1) {
				var namest = cell.parentColumn.index + 1;
				var nameend = namest + cell.columnSpan - 1;
				cells[i].xmlAttributes.add('namest', 'col' + namest);
				cells[i].xmlAttributes.add('nameend', 'col' + nameend);
			}
			//추가 속성 by iM - 20.07.30
			cells[i].xmlAttributes.add('Name', cell.name + '');
			cells[i].xmlAttributes.add('RowSpan', cell.rowSpan + '');
			cells[i].xmlAttributes.add('ColumnSpan', cell.columnSpan + '');
			cells[i].xmlAttributes.add('AppliedCellStyle', cell.appliedCellStyle.name + '');
			if (cell.firstBaselineOffset == "1296135023") {
				cellfBaseline = "ASCENT_OFFSET";
			} if (cell.firstBaselineOffset == "1296255087") {
				cellfBaseline = "CAP_HEIGHT";
			} if (cell.firstBaselineOffset == "1296386159") {
				cellfBaseline = "EMBOX_HEIGHT";
			} if (cell.firstBaselineOffset == "1313228911") {
				cellfBaseline = "FIXED_HEIGHT";
			} if (cell.firstBaselineOffset == "1296852079") {
				cellfBaseline = "LEADING_OFFSET";
			} if (cell.firstBaselineOffset == "1299728495") {
				cellfBaseline = "X_HEIGHT";
			}
			cells[i].xmlAttributes.add('FirstBaselineOffset', cellfBaseline + '');
			if (cell.verticalJustification == "1953460256") {
				cellVerticajust = "TOP_ALIGN";
			} if (cell.verticalJustification == "1667591796") {
				cellVerticajust = "CENTER_ALIGN";
			} if (cell.verticalJustification == "1651471469") {
				cellVerticajust = "BOTTOM_ALIGN";
			} if (cell.verticalJustification == "1785951334") {
				cellVerticajust = "JUSTIFY_ALIGN";
			}
			cells[i].xmlAttributes.add('VerticalJustification', cellVerticajust + '');
			cells[i].xmlAttributes.add('LeftEdgeStrokeWeight', cell.leftEdgeStrokeWeight + '');
			cells[i].xmlAttributes.add('RightEdgeStrokeWeight', cell.rightEdgeStrokeWeight + '');
			cells[i].xmlAttributes.add('TopEdgeStrokeWeight', cell.topEdgeStrokeWeight + '');
			cells[i].xmlAttributes.add('BottomEdgeStrokeWeight', cell.bottomEdgeStrokeWeight + '');
			cells[i].xmlAttributes.add('FillColor', cell.fillColor.name);
			cells[i].xmlAttributes.add('FillTint', cell.fillTint + '');
			//추가 속성 by iM - 20.07.30
		}
	}
	AddImageSize(myDoc)
	function AddImageSize(myDoc) {
		try {
			var result = true;

			var allImages = myDoc.allGraphics;
			var imageCount = GetImageCountWithoutMaster(myDoc);
			for (var i = 0; i < imageCount; i++) {
				try {
					allImages[i].associatedXMLElement.xmlAttributes.itemByName('href').value = decodeURI(allImages[i].itemLink.filePath);
					// allImages[i].itemLink.filePath
					allImages[i].associatedXMLElement.xmlAttributes.add('xScale', allImages[i].absoluteHorizontalScale.toString());
					allImages[i].associatedXMLElement.xmlAttributes.add('yScale', allImages[i].absoluteVerticalScale.toString());
					//이미지 오브젝트의 offset 값을 넣어주는 함수 by iM - 20.07.31
					var anchorPosition = "";
					if (allImages[i].parent.anchoredObjectSettings.anchoredPosition == "1095716961") {
						anchorPosition = "ABOVE_LINE";
					} if (allImages[i].parent.anchoredObjectSettings.anchoredPosition == "1097814113") {
						anchorPosition = "ANCHORED";
					} if (allImages[i].parent.anchoredObjectSettings.anchoredPosition == "1095716969") {
						anchorPosition = "INLINE_POSITION";
					}
					allImages[i].associatedXMLElement.xmlAttributes.add('AnchoredPosition', anchorPosition);

					var xoffset = allImages[i].parent.anchoredObjectSettings.anchorXoffset * 2.835;
					var yoffset = allImages[i].parent.anchoredObjectSettings.anchorYoffset * 2.835;
					
					allImages[i].associatedXMLElement.xmlAttributes.add('xOffset', xoffset + '');
					allImages[i].associatedXMLElement.xmlAttributes.add('yOffset', yoffset + '');
					allImages[i].associatedXMLElement.xmlAttributes.add('appliedObjectStyle', allImages[i].parent.appliedObjectStyle.name);
					var cropOpt = allImages[i].parent.frameFittingOptions;

					if (cropOpt.leftCrop != 0) {
						allImages[i].associatedXMLElement.xmlAttributes.add('LeftCrop', cropOpt.leftCrop * 2.835 + '');
					} if (cropOpt.rightCrop != 0) {
						allImages[i].associatedXMLElement.xmlAttributes.add('RightCrop', cropOpt.rightCrop * 2.835 + '');
					} if (cropOpt.topCrop != 0) {
						allImages[i].associatedXMLElement.xmlAttributes.add('TopCrop', cropOpt.topCrop * 2.835 + '');
					} if (cropOpt.bottomCrop != 0 ) {
						allImages[i].associatedXMLElement.xmlAttributes.add('BottomCrop', cropOpt.bottomCrop * 2.835 + '');
					}
					//이미지 오브젝트의 offset 값을 넣어주는 함수 by iM - 20.07.31
				} catch (ex) {
					ex;
				}
			}
			return result;
		} catch (ex) {
			alert(ex);
		}
	}

	function GetImageCountWithoutMaster(myDoc) {
		try {
			var myMaster = myDoc.masterSpreads;
			var myMasterNumber = myMaster.count();
			var masterImages = 0;
			for (var i = 0 ; i < myMasterNumber ; i++) {
				imageNumber = myMaster[i].allGraphics;
				masterImages += imageNumber.length;
			}

			return myDoc.allGraphics.length - masterImages;
		} catch (ex) {
			alert(ex);
		}
	}

	//XML 출력
	var script = app.activeScript;
	var myScriptFolderPath = script.path;
	var indentXML = new File(myScriptFolderPath  + '/indent_XML.xsl');
	var myName = myDoc.name.split(".indd").join(".xml");
	app.scriptPreferences.userInteractionLevel=UserInteractionLevels.neverInteract;
	var myInxfile = new File(myDoc.filePath.absoluteURI + "/Translate/" + myDoc.name.split(".indd")[0]+".xml");
	if (indentXML.exists) {
		myDoc.xmlExportPreferences.allowTransform = true;
		myDoc.xmlExportPreferences.transformFilename = indentXML.fsName;
		// myDoc.exportFile(ExportFormat.xml, FileOutput);
		myDoc.exportFile(ExportFormat.XML, myInxfile);
	} else {
		throw new Error('Reference 폴더에 indent_XML.xsl 파일이 없습니다. 프로그램이 손상되었습니다.');
	}
	
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