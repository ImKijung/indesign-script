var doc = app.activeDocument;
var myPStyle = doc.allParagraphStyles;
var myCStyle = doc.allCharacterStyles;
doc.xmlPreferences.defaultCellTagName = "Cell";
doc.xmlPreferences.defaultImageTagName = "Image";
doc.xmlPreferences.defaultStoryTagName = "Story";
doc.xmlPreferences.defaultTableTagName = "Table";

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
try {
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
	for (var y=0; y < gStory.length; y++) {
		var myGraphics = gStory[y].allGraphics;
		for (x=0; x < myGraphics.length; x++){
			myGraphics[x].autoTag();
		}
	}
} catch(e) {}

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

// 강제 줄바꿈 해제
app.findGrepPreferences = app.findChangeGrepOptions = null;
app.changeGrepPreferences = app.findChangeGrepOptions = null;
	
	app.findGrepPreferences.properties = {
		findWhat : "\\n",
	}
	app.findChangeGrepOptions.properties = {
		includeFootnotes:false,
		includeHiddenLayers:false,
		includeLockedLayersForFind:false,
		includeLockedStoriesForFind:false,
		includeMasterPages:false,
	}
	app.changeGrepPreferences.changeTo = "",
doc.changeGrep();
app.findGrepPreferences = app.findChangeGrepOptions = null;
app.changeGrepPreferences = app.findChangeGrepOptions = null;

//문자스타일에 Tag 적용하기
//C_Image 해제하기
app.findTextPreferences = app.changeTextPreferences = NothingEnum.nothing;
app.findTextPreferences.appliedCharacterStyle = "C_image";
app.changeTextPreferences.appliedCharacterStyle = app.characterStyles.item("$ID/[None]");

doc.changeText();

for (c=1; c<myCStyle.length; c++) {
     var myCStyleName = myCStyle[c];
     app.findTextPreferences = app.changeTextPreferences = NothingEnum.nothing;
     app.findTextPreferences.appliedCharacterStyle = myCStyleName;
     found = doc.findText();
     for (d=1; d<found.length; d++) {
          //var founds = found.length;
          text = found[d];
          found[d].select();
          selection = app.selection[0];
          doc.xmlElements[0].xmlElements.add('span', selection).xmlAttributes.add('style', String(myCStyleName.name));
     }
     app.findTextPreferences = app.changeTextPreferences = NothingEnum.nothing;
}

//paraXML();
//단락스타일에 Tag 적용하기
function paraXML() {
	
	var doc = app.properties.activeDocument,
	fgp = app.findGrepPreferences.properties,
	fcgo = app.findChangeGrepOptions.properties,
	found, n = 0, text, pTag, xe;
	
	//Exit if no documents open
	if ( !doc ) {
		alert("You need an open document" );
		return;
	}
	
	//Setting F/R Grep
	app.findGrepPreferences = app.findChangeGrepOptions = null;
	
	app.findGrepPreferences.properties = {
		findWhat : "^.+",
	}
	app.findChangeGrepOptions.properties = {
		includeFootnotes:false,
		includeHiddenLayers:false,
		includeLockedLayersForFind:false,
		includeLockedStoriesForFind:false,
		includeMasterPages:false,
	}
	
	//Adding "p" tag if needed
	pTag = doc.xmlTags.itemByName("para");
	!pTag.isValid && pTag = doc.xmlTags.add({name:"para"});
	
	//Getting paragraphs occurences
	found = doc.findGrep();
	n = found.length;
	while ( n-- ) {
		text = found[n];
          text.select();
          sel = app.selection[0];
          myPStyleName = sel.paragraphs[0].appliedParagraphStyle;
		xe = text.associatedXMLElements;
		//Adding "p" tags with ids if needed
		if ( !xe.length || xe[0].markupTag.name!="para") { 
			var myXML = doc.xmlElements[0].xmlElements.add( pTag, text );
               myXML.xmlAttributes.add('id', guid () );
               myXML.xmlAttributes.add('style', String(myPStyleName.name));
               //xmlRoot.xmlElements.add('para', text).xmlAttributes.add('name', String(myPStyleName.name));
		}
	}
	
	//Reverting initial F/R settings
	app.findGrepPreferences.properties = fgp;
	app.findChangeGrepOptions.properties = fcgo;
}

//Returns unique ID
function guid() {
  function s4() {
    return Math.floor((1 + Math.random()) * 0x10000)
      .toString(16)
      .substring(1);
  }
  return s4() + s4() + '-' + s4() + '-' + s4() + '-' + s4() + '-' + s4() + s4() + s4();
}

//Table의 스타일 네임 속성 적용하기
var rootXML = doc.xmlElements[0];
var tbElement = rootXML.evaluateXPathExpression (".//Table");
for (var d=0; d<tbElement.length; d++) {
	app.select(tbElement[d]);
	sel = app.selection[0];
	myTBtyleName = sel.tables[0].appliedTableStyle;
	tbElement[d].xmlAttributes.add('style', String(myTBtyleName.name));
}
//Cell의 스타일 네임 속성 적용하기
var cellElement = rootXML.evaluateXPathExpression (".//Cell");
for (var e=0; e<cellElement.length; e++) {
	app.select(cellElement[e]);
	sel = app.selection[0];
	myCelltyleName = sel.cells[0].appliedCellStyle;
	cellElement[e].xmlAttributes.add('style', String(myCelltyleName.name.split("없음").join("None")));
}

//Hyperlink의 Tag 적용하기
var myRef = doc.crossReferenceSources.everyItem().getElements();
for (var f=0; f<myRef.length; f++) {
	app.select(myRef[f].sourceText);
	selection = app.selection[0];
	doc.xmlElements[0].xmlElements.add('xref', selection).xmlAttributes.add('href', String(myRef[f].name));
}

//Destination의 Tag 적용하기
link = doc.hyperlinkTextDestinations;
for (var k=0 ; k<link.length; k++) {
	var linkName = link[k].name;
	var xxx = link[k].destinationText.paragraphs[0];
	xxx.select();
	sel = app.selection[0];
	xe = xxx.associatedXMLElements;
	var myXML = doc.xmlElements[0].xmlElements.add( "anchor", xxx );
		myXML.xmlAttributes.add('x', String(linkName) );
}

//XML 출력
// var myName = doc.name.split(".indd").join(".xml");
// // var myFolder = Folder.selectDialog("Select the folder", "");
// app.scriptPreferences.userInteractionLevel=UserInteractionLevels.neverInteract;    
// var myInxfile = new File(doc.filePath.absoluteURI + "/" + doc.name.split(".indd")[0]+".xml");
// doc.exportFile(ExportFormat.XML, myInxfile);
// alert("XML Export 완료")
// doc.close(SaveOptions.NO);