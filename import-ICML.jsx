//Save the current application setting.
var currentAppSettings = {checkLinksAtOpen: app.linkingPreferences.checkLinksAtOpen};

//Set the value to false to prevent the dialog from showing.
app.linkingPreferences.checkLinksAtOpen = false;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;

var doc = app.activeDocument;
myFile = File.openDialog("Choose an ICML file:", false);
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
doc.pages.item(2).textFrames[0].place(myFile);

var myFileName = myFile.name.substr(0,myFile.name.lastIndexOf("."));
var linkICML = myFileName + ".icml";
doc.links.item(linkICML).unlink();
myFile.close();

//Set the value back to its original value.
app.linkingPreferences.checkLinksAtOpen = currentAppSettings.checkLinksAtOpen;

for (var i = 0; i < doc.allGraphics.length; i++)
{
    var target = doc.allGraphics[i].parent;
    target.fit(FitOptions.frameToContent);
}

//목차 만들기
doc.pages.item(1).textFrames[1].select();
app.scriptMenuActions.itemByID(71442).invoke();

var tocpStyle = doc.paragraphStyles.itemByName('TOC2');
var toccStyle = doc.characterStyles.itemByName('C_TryNow-Invisible');

app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;

app.findGrepPreferences.findWhat = "[^>]*";  
app.findGrepPreferences.appliedParagraphStyle = tocpStyle;
app.findGrepPreferences.appliedCharacterStyle = toccStyle;
app.changeGrepPreferences.changeTo = "\\r";  
app.changeGrepPreferences.appliedParagraphStyle = tocpStyle;
app.changeGrepPreferences.appliedCharacterStyle = null;

doc.changeGrep();

app.findTextPreferences = NothingEnum.nothing;
app.changeTextPreferences = NothingEnum.nothing;


//표지 만들기
default xml namespace = "http://www.idiominc.com/opentopic/vars";
var langcode = doc.textVariables.item("langcode").variableOptions.contents;
var varfile = "C:\\DITA-OT\\plugins\\com.company.custpdf.html5\\cfg\\common\\vars\\" + langcode + ".xml";
var f = new File(varfile);

if (f != null)
{
    f.open("r");
    var varsXML = new XML(f.read());
    f.close();

    var frontmatter_title = varsXML.variable.(@id == "frontmatter-title").toString();
    var frontmatter_thankyou = varsXML.variable.(@id == "frontmatter-thankyou").toString();
    var frontmatter_inform_register = varsXML.variable.(@id == "frontmatter-inform-register").toString();
    var frontmatter_url = varsXML.variable.(@id == "frontmatter-url").toString();
    var frontmatter_model = varsXML.variable.(@id == "frontmatter-model").toString();
    var frontmatter_serialno = varsXML.variable.(@id == "frontmatter-serialno").toString();
    var frontmatter_inform_menuscreen = varsXML.variable.(@id == "frontmatter-inform-menuscreen").toString();
    var frontmatter_learnmenuscreen = varsXML.variable.(@id == "frontmatter-learnmenuscreen").toString();
    var frontmatter_link = varsXML.variable.(@id == "frontmatter-link").toString();
    var frontmatter_open_quot = varsXML.variable.(@id == "frontmatter-open-quot").toString();
    var frontmatter_close_quot = varsXML.variable.(@id == "frontmatter-close-quot").toString();
    var frontmatter_toc_contents = varsXML.variable.(@id == "frontmatter-toc-contents").toString();
    // doc.textVariables.item("Contents").variableOptions.contents = frontmatter_toc_contents;

    doc.pages.item(0).textFrames[0].select();
    app.selection[0].contents = frontmatter_title;
    
    var tempStr = frontmatter_thankyou + "\r" +
                  frontmatter_inform_register +  "\r\r" +
                  frontmatter_url + "\r\r" +
                  frontmatter_model + " __________________ " + frontmatter_serialno + " __________________ " + "\r\r" +
                  frontmatter_inform_menuscreen;

    switch ( langcode ) {
        case "fr":
            tempStr += "\r" + frontmatter_link + " " + frontmatter_open_quot + " " + frontmatter_learnmenuscreen + " " + frontmatter_close_quot;
            break;
        default:
            tempStr += "\r" + frontmatter_open_quot + frontmatter_learnmenuscreen + frontmatter_close_quot + " " + frontmatter_link;
            break;
    }
                   
    doc.pages.item(0).textFrames[1].select();
    app.selection[0].contents = tempStr;
}

inlineStyling();
create_hyperlink(frontmatter_learnmenuscreen);
app.selection = null;

function inlineStyling() {
    //문자 스타일 적용 스크립트
    var psName1 = doc.paragraphStyles.itemByName('Cover_Description');
    var csName1 = doc.characterStyles.itemByName('TOC_SingleLine');

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences = app.changeGrepPreferences = null;
    app.findTextPreferences.findWhat = "__________________";
    app.findTextPreferences.appliedParagraphStyle = psName1;
    app.changeTextPreferences.changeTo = "__________________";
    app.changeTextPreferences.appliedParagraphStyle = psName1;
    app.changeTextPreferences.appliedCharacterStyle = csName1;
    app.activeDocument.changeText();

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    //텍스트 앵커 만들기
    app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.nothing;
    app.findGrepPreferences.findWhat = ".+";
    app.findGrepPreferences.appliedParagraphStyle = doc.paragraphStyles.item("Subchapter");
    var finds = doc.findGrep();

    var destCounter = 0;

    for ( var j = finds.length-1; j >= 0; j-- ) {
        var found = finds[j];
        try {
            if (!doc.hyperlinkTextDestinations.itemByName(found.contents).isValid) {
                var hypTextDest = doc.hyperlinkTextDestinations.add(found);
                hypTextDest.name = found.contents;
                destCounter++;
            }
        }
        catch(e) {}
    }
}

//하이퍼링크 만들기
function create_hyperlink(learnmenuscreen) {
	var doc = app.activeDocument;
	var xRefForm = app.activeDocument.crossReferenceFormats.item("단락 텍스트");
	var target_text = doc.hyperlinkTextDestinations.item(learnmenuscreen).destinationText;
	var destination = doc.paragraphDestinations.add (target_text);

    app.findTextPreferences = NothingEnum.nothing;
    app.changeTextPreferences = NothingEnum.nothing;

    app.findTextPreferences = app.changeGrepPreferences = null;
	app.findTextPreferences.findWhat = learnmenuscreen;
	app.findTextPreferences.appliedParagraphStyle = "Cover_Description";

    var myFinds = app.activeDocument.findText();
	var spec_counter = 0;

    for (var i=0;i<myFinds.length;i++) {
		var sText = myFinds[1].select();
	}

    var source_text = app.selection[0].texts[0];
    source_text.appliedCharacterStyle = "C_Font-Underline";
	var source = doc.crossReferenceSources.add (source_text, xRefForm);
    doc.hyperlinks.add (source, destination);
}