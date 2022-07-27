var doc = app.activeDocument;
var tocStyleName = "TOC";
var myTOCStyle = doc.tocStyles.itemByName( tocStyleName );
if( myTOCStyle.isValid ) {
    doc.createTOC( myTOCStyle , true ); // Second argument: Update existing TOC
};