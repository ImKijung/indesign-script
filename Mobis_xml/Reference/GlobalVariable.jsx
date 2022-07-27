var errorMsgs = new Array();
var hasError = false;
var isStopped = false;
var CloseWithoutMsg = 'CloseWithoutMsg';

var imageSize = 100;
var folderPath = null;
var linked = [];

var askSave = true;
var askBook = true;

var currentFolder = new File($.fileName).parent;
var referenceFolder = null;
if (currentFolder.name == 'Reference')
    referenceFolder = currentFolder;
else
    referenceFolder = new Folder(currentFolder.fsName + '\\Reference');
var indentXML = new File(referenceFolder.fsName + '\\indent_XML.xsl');
var inddFile = new File(referenceFolder.fsName + '\\ImageCharacterStyle.indd');
var DBxmlFile = new File(referenceFolder.fsName + '\\DB.xml');
DBxmlFile.open('r');
var DBxml = new XML(DBxmlFile.read());
DBxmlFile.close();
var imageStyleName = DBxml.xpath("/Root/CharacterStyles/*[@type='CharImage'][1]").attribute('name').toString();
var appImageStyleName = DBxml.xpath("/Root/CharacterStyles/*[@type='CharImage'][1]").attribute('globalName').toString();
var floatIdStyles = DBxml.xpath("/Root/ParagraphStyles/*[@type='floatId']");
var TOCremoveStyles = DBxml.xpath("/Root/CharacterStyles/*[@type='TOC']");
//imageStyleName = 'C_Image'
var styledImageTables = DBxml.xpath("/Root/TableStyles/*[@type='useImageCharacterStyle']");
var higthlightStyles = DBxml.xpath("//*[@type='highlight']");

var THICK = '5mm';//밑줄 두께
var ChangedTHICK = '5mm';//밑줄 두께
var SHIFT = '3pt';//베이스라인 높이

var progresspanel = new Window("window", "실행중");
with (progresspanel) { progresspanel.progressbar = add("progressbar", [12, 12, 400, 24], 0, 100); }
progressTitle = "실행 중...";



function SetGlobalVariableInitialize() {
    errorMsgs = new Array();
    hasError = false;
    isStopped = false;
    imageSize = 100;
    folderPath = null;
    linked = [];
    askSave = true;
    askBook = true;
    progressTitle = "실행 중...";
    progresspanel.text = progressTitle;
}
