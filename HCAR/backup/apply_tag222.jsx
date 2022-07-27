var doc = app.activeDocument;
var myPStyle = doc.allParagraphStyles;
var myCStyle = doc.allCharacterStyles;
doc.xmlPreferences.defaultCellTagName = "Cell";
doc.xmlPreferences.defaultImageTagName = "Image";
doc.xmlPreferences.defaultStoryTagName = "Story";
doc.xmlPreferences.defaultTableTagName = "Table";


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
