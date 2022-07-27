//MAIN ROUTINE
var main = function() {
	
	var doc = app.properties.activeDocument,
	fgp = app.findTextPreferences.properties,
	fcgo = app.findChangeTextOptions.properties,
	found, n = 0, text, pTag, xe;
	
	//Exit if no documents open
	if ( !doc ) {
		alert("You need an open document" );
		return;
	}
	
	//Setting F/R Text
	app.findTextPreferences = app.findChangeTextOptions = null;
	
	app.findTextPreferences.properties = {
		findWhat : "<0016>",
	}
	app.findChangeTextOptions.properties = {
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
	found = doc.findText();
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
               myXML.xmlAttributes.add('name', String(myPStyleName.name));
               //xmlRoot.xmlElements.add('para', text).xmlAttributes.add('name', String(myPStyleName.name));
		}
	}
	
	//Reverting initial F/R settings
	app.findTextPreferences.properties = fgp;
	app.findChangeTextOptions.properties = fcgo;
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

var u;

//Run
main();