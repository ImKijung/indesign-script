var doc = app.activeDocument;
var myPStyle = doc.allParagraphStyles;
var myCStyle = doc.allCharacterStyles;

var rootXML = doc.xmlElements[0];
var tbElement = rootXML.evaluateXPathExpression (".//para");

for (c=1; c<myCStyle.length; c++) {
     var myCStyleName = myCStyle[c];
     app.findTextPreferences = app.changeTextPreferences = NothingEnum.nothing;
     app.findTextPreferences.appliedCharacterStyle = myCStyleName;
     found = doc.findText();
     for (d=0; d<found.length; d++) {
          //var founds = found.length;
          text = found[d];
          found[d].select();
          selection = app.selection[0];
          doc.xmlElements[0].xmlElements.add('char', selection).xmlAttributes.add('style', String(myCStyleName.name));
     }
     app.findTextPreferences = app.changeTextPreferences = NothingEnum.nothing;
}