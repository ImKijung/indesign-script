
var doc = app.activeDocument;
var tocpStyle = doc.paragraphStyles.itemByName('TOC2');
var toccStyle = doc.characterStyles.itemByName('C_TryNow-Invisible');

app.findGrepPreferences.findWhat = "[^>]*";  
app.findGrepPreferences.appliedParagraphStyle = tocpStyle;
app.findGrepPreferences.appliedCharacterStyle = toccStyle;
app.changeGrepPreferences.changeTo = "\\r";  
app.changeGrepPreferences.appliedParagraphStyle = tocpStyle;
app.changeGrepPreferences.appliedCharacterStyle = null;

doc.changeGrep();

//app.findTextPreferences = NothingEnum.nothing;
//app.changeTextPreferences = NothingEnum.nothing;