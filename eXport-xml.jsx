/*
  InDesignデータから，テキスト構造をXMLとして書き出す
  Version 1.1.0
  Changes
  1.0.0: initial version
  1.0.1: ルビをサポート (2015/10/1)
  1.0.2: 複数ファイルをいっぺんに処理したとき，XMLを一つにまとめるよう変更
  1.0.3: paragraph.parentTextFrames[0]がTextPathだったときの処理を追加
  1.1.0: 大幅に書き直し。階層的にオブジェクトを取り出せるよう修正。特色を使っているスタイルは，特色情報も取り出せるように変更。
*/
#target indesign

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var CHARSTYLE_NONE = "[なし]";

function dumpObj(obj) {
    $.writeln("----");
    $.writeln(obj.toString());
    for(var prop in obj) {
	try {
	    $.writeln("name: " + prop + "; value: " + obj[prop]);
	}
	catch(e) {
	    $.writeln("name: " + prop + "; cannot access this property.");
	}
    }
}

function dumpObjToFile(obj, filename) {
    var file = new File(filename);
    if(file && file.open("w")) {
	file.encoding = "utf-8";
	file.writeln(obj.toString());
	for(var prop in obj) {
	    try {
		file.writeln("name: " + prop + "; value: " + obj[prop]);
	    }
	    catch(e) {
		file.writeln("name: " + prop + "; cannot access this property.");
	    }
	}
	file.close();
    }    
}

function selectIdFiles(folder) {
    var idFiles = [];
    var files = folder.getFiles("*");
    for(var i=0; i<files.length; ++i) {
	var f = files[i];
	if(f instanceof Folder) {
	    idFiles = idFiles.concat(selectIdFiles(f));
	}
	else if(f.name.match(/\.indd$/) && f.name.charAt(0) != ".") {
	    // ファイル名の先頭が「.」ではじまるファイルは除外
	    idFiles.push(f);
	}
    }
    return idFiles.sort();
}

function PdfVisibleLayers(pdf) {
    var layers = pdf.graphicLayerOptions.graphicLayers;
    var visibleLayers = "";
    for(var i=0; i<layers.length; ++i) {
	var layer = layers[i];
	if(layer.currentVisibility) {
	    visibleLayers += layer.name;
	}
    }
    return visibleLayers;
}

function GetStoryFirstTextFrame(story) {
    if(story && story.isValid) {
	var para = story.paragraphs[0];
	if(para && para.isValid) {
	    var tf = para.parentTextFrames[0];
	    while(tf instanceof TextPath) {
		tf = tf.parent;
	    }
	    return tf;
	}
    }
    return null;
}

function GetStoryInfo(story) {
    var tf = GetStoryFirstTextFrame(story);
    var info;
    if(tf && tf.isValid && tf.parentPage) {
	info = {
	    story: story,
	    page: tf.parentPage.name,
	    top: tf.visibleBounds[0],
	    left: tf.visibleBounds[1]
	}
    }
    return info;
}

function Analyzer() {
}

Analyzer.prototype.getChild = function(item) {
    if(item.allPageItems.length > 0) {
	return this.getChild(item.allPageItems[0]);
    }
    else {
	return item;
    }
}

Analyzer.prototype.analyzeDoc = function(doc) {
    this.doc = doc;

    var pageBinding = doc.documentPreferences.pageBinding;
    var stories = [];
    var docXml = new XML("<file/>");
    docXml.@name = doc.name;
    var pagesXml = new XML("<pages/>");
    var pageXml;
    var prevPage;
    var page;
    var story;
    var storyInfo;
    var storyIndex;

    for(storyIndex=0; storyIndex<doc.stories.length; ++storyIndex) {
	story = doc.stories[storyIndex];
	if(story && story.isValid) {
	    storyInfo = GetStoryInfo(story);
	    if(storyInfo) {
		stories.push(storyInfo);
	    }
	}
    }
    stories.sort(function(a, b) {
	var pa = a.page;
	var pb = b.page;
	if(pa != pb) {
	    var pan = Number(pa);
	    var pbn = Number(pb);
	    if(isNaN(pan) && isNaN(pbn)) {
		// a.page, b.pageともに文字列
		if(String(pa) > String(pb)) {
		    return 1;
		}
		else {
		    return -1;
		}
	    }
	    else if(isNaN(pan)) {
		return 1;
	    }
	    else if(isNaN(pbn)) {
		return -1;
	    }
	    else {
		// a.page, b.pageともに数値
		return pan - pbn;
	    }
	}
	else {
	    var ta = a.top;
	    var tb = b.top;
	    var la = a.left;
	    var lb = b.left;
	    if(pageBinding == PageBindingOptions.RIGHT_TO_LEFT) {
		// 右綴じ
		if(Math.abs(ta - tb) > 10) {
		    return ta - tb;
		}
		else {
		    return lb - la;
		}
	    }
	    else {
		// 左綴じ
		if(Math.abs(la - lb) > 10) {
		    return la - lb;
		}
		else {
		    return ta - tb;
		}
	    }
	}
	return 0;
    });

    for(storyIndex=0; storyIndex<stories.length; ++storyIndex) {
	storyInfo = stories[storyIndex];
	story = storyInfo.story;
	if(story && story.isValid) {
	    page = storyInfo.page;
	    if(!pageXml || page != prevPage) {
		prevPage = page;
		pageXml = new XML("<page/>");
		pageXml.@number = page;
		pagesXml.appendChild(pageXml);
	    }
	    pageXml.appendChild(this.analyzeStory(story));
	}
    }
    
    docXml.appendChild(pagesXml);
    return docXml;
}

Analyzer.prototype.analyzeStory = function(story) {
    var paragraph;
    var paraIndex;
    var textFrame;
    var textFrameParent;
    var storyXml = new XML("<story/>");

    for(paraIndex=0; paraIndex<story.paragraphs.length; ++paraIndex) {
	paragraph = story.paragraphs[paraIndex];
	textFrame = paragraph.parentTextFrames[0];
	while(textFrame instanceof TextPath) {
	    textFrame = textFrame.parent;
	}
	if(textFrame && textFrame.isValid) {
	    textFrameParent = textFrame.parent;
	    if(!(textFrameParent instanceof(Character)) &&  paragraph.isValid) {
		storyXml.appendChild(this.analyzeParagraph(paragraph));
	    }
	}
    }

    if(storyXml.length() == 0) {
	storyXml = null;
    }
    
    return storyXml;
}

Analyzer.prototype.analyzeParagraph = function(paragraph) {
    var paraXml = new XML("<p/>");
    var contentXml = paraXml;
    var charIndex;
    var ch;
    var ch_s;
    var chStyle = null;
    var text = "";
    var anchorObject;
    var specialCharXml;
    var tableIndex;
    var rubyFlag;
    var rubyString;
    var rubyXml;
    var rubyStringXml;
    var spotColorName;
    var rubyColorName;
    var charCount = paragraph.characters.length;
    paraXml.@style = paragraph.appliedParagraphStyle.name;

    if(paragraph.tables.length > 0) {
	for(tableIndex=0; tableIndex<paragraph.tables.length; ++tableIndex) {
	    contentXml.appendChild(this.analyzeTable(paragraph.tables[tableIndex]));
	}
    }
    else {
	for(charIndex=0; charIndex<charCount; ++charIndex) {
	    ch = paragraph.characters[charIndex];
	    if(ch.appliedCharacterStyle.isValid &&
	       ch.appliedCharacterStyle.name != chStyle) {
		// 文字スタイルが変更された
		if(rubyFlag) {
		    rubyXml = new XML("<ruby/>");
		    rubyXml.appendChild(text);
		    text = "";
		    rubyStringXml = new XML("<rt/>");
		    if(rubyColorName) {
			rubyStringXml.@color = rubyColorName;
			rubyColorName = null;
		    }
		    rubyStringXml.appendChild(rubyString);
		    rubyXml.appendChild(rubyStringXml);
		    rubyString = "";
		    rubyFlag = false;
		    contentXml.appendChild(rubyXml);
		}
		else {
		    chStyle = ch.appliedCharacterStyle.name;
		    contentXml.appendChild(text);
		    text = "";
		    if(chStyle == CHARSTYLE_NONE) { // 文字スタイル無し
			contentXml = paraXml;
		    }
		    else { // 新規文字スタイル
			contentXml = new XML("<c/>");
			contentXml.@style = chStyle;
			paraXml.appendChild(contentXml);
		    }
		}
	    }
	    if(ch.allPageItems.length > 0) {
		// アンカーオブジェクト
		contentXml.appendChild(text);
		text = "";
		anchorObject = this.getChild(ch);
		if(anchorObject instanceof(TextFrame)) {
		    contentXml.appendChild(this.analyzeAnchorText(anchorObject));
		}
		else if(anchorObject instanceof(Graphic)
			|| anchorObject instanceof(PDF)) {
		    contentXml.appendChild(this.analyzeAnchorGraphic(anchorObject));
		}
	    }
	    else {
		// 文字
		ch_s = ch.contents.toString();
		
		if(spotColorName != ch.fillColor.name) {
		    if(ch.fillColor instanceof(Color)
		       && ch.fillColor.model == ColorModel.SPOT) {
			// 特色
			contentXml.@color = ch.fillColor.name;
			spotColorName = ch.fillColor.name;
		    }
		    else {
			spotColorName = null;
		    }
		}
		
		if(!rubyFlag && ch.rubyFlag) {
		    // ルビ開始
		    contentXml.appendChild(text);
		    text = "";
		    rubyFlag = true;
		    rubyString = ch.rubyString;
		    if(ch.rubyFill instanceof(Color)
		       && ch.rubyFill.model == ColorModel.SPOT) {
			rubyColorName = ch.rubyFill.name;
		    }
		    else {
			rubyColorName = null;
		    }
		}
		else if(rubyFlag && !ch.rubyFlag) {
		    // ルビ終了
		    rubyXml = new XML("<ruby/>");
		    rubyXml.appendChild(text);
		    text = "";
		    rubyStringXml = new XML("<rt/>");
		    if(rubyColorName) {
			rubyStringXml.@color = rubyColorName;
			rubyColorName = null;
		    }
		    rubyStringXml.appendChild(rubyString);
		    rubyXml.appendChild(rubyStringXml);
		    rubyString = "";
		    rubyFlag = false;
		    contentXml.appendChild(rubyXml);
		}
		else if(rubyFlag && ch.rubyFlag && rubyString != ch.rubyString) {
		    // ルビ切り替え
		    rubyXml = new XML("<ruby/>");
		    rubyXml.appendChild(text);
		    text = "";
		    rubyStringXml = new XML("<rt/>");
		    if(rubyColorName) {
			rubyStringXml.@color = rubyColorName;
			rubyColorName = null;
		    }
		    rubyXml.appendChild(rubyStringXml);
		    rubyStringXml.appendChild(rubyString);
		    rubyString = ch.rubyString;
		    contentXml.appendChild(rubyXml);
		}
		
		if(ch_s.length > 1) {
		    // 特殊文字
		    switch(ch_s) {
		    case "SINGLE_RIGHT_QUOTE":
			text += "’";
			break;
		    case "SINGLE_LEFT_QUOTE":
			text += "‘";
			break;
		    case "DOUBLE_RIGHT_QUOTE":
			text += "”";
			break;
		    case "DOUBLE_LEFT_QUOTE":
			text += "“";
			break;
		    default:
			contentXml.appendChild(text);
			text = "";
			specialCharXml = new XML("<specialChar/>");
			specialCharXml.@name = ch_s;
			contentXml.appendChild(specialCharXml);
		    }
		}
		else if(ch_s.match(/\t/)) {
		    contentXml.appendChild(text);
		    text = "";
		    specialCharXml = new XML("<specialChar/>");
		    specialCharXml.@name = "TAB";
		    contentXml.appendChild(specialCharXml);
		}
		else if(ch_s.charCodeAt(0) < 0x20) {
		    // 制御文字（スペースに置き換え）
		    text += " ";
		}
		else {
		    text += ch_s;
		}
	    }
	}

	if(rubyFlag) {
	    rubyXml = new XML("<ruby/>");
	    rubyXml.appendChild(text);
	    text = "";
	    rubyStringXml = new XML("<rt/>");
	    rubyStringXml.appendChild(rubyString);
	    rubyXml.appendChild(rubyStringXml);
	    rubyString = "";
	    rubyFlag = false;
	    contentXml.appendChild(rubyXml);
	}
	else {
	    contentXml.appendChild(text);
	}
    }

    return paraXml;
}

Analyzer.prototype.analyzeTable = function(table) {
    var tableXml = new XML("<table/>");
    var rowXml;
    var cellXml;
    var rowIndex;
    var cellIndex;
    var paraIndex;
    var row;
    var cell;
    var paragraph;

    for(rowIndex=0; rowIndex<table.rows.length; ++rowIndex) {
	row = table.rows[rowIndex];
	rowXml = new XML("<tr/>");
	for(cellIndex=0; cellIndex<row.cells.length; ++cellIndex) {
	    cell = row.cells[cellIndex];
	    cellXml = new XML("<td/>");
	    for(paraIndex=0; paraIndex<cell.paragraphs.length; ++paraIndex) {
		paragraph = cell.paragraphs[paraIndex];
		cellXml.appendChild(this.analyzeParagraph(paragraph));
	    }
	    rowXml.appendChild(cellXml);
	}

	tableXml.appendChild(rowXml);
    }

    return tableXml;
}

Analyzer.prototype.analyzeAnchorGraphic = function(anchorGraphic) {
    var xml = new XML("<anchorGraphic/>");
    xml.@name = anchorGraphic.itemLink.name;
    xml.@layer = PdfVisibleLayers(anchorGraphic);
    return xml;
}

Analyzer.prototype.analyzeAnchorText = function(anchorText) {
    var xml = new XML("<anchorText/>");
    var paraIndex;
    var paragraph;
    for(paraIndex=0; paraIndex<anchorText.paragraphs.length; ++paraIndex) {
	paragraph = anchorText.paragraphs[paraIndex];
	xml.appendChild(this.analyzeParagraph(paragraph));
    }

    if(xml.length() == 0) {
	xml = null;
    }
    
    return xml;
}

var getSaveFile = function() {
    var file = new File("export_text.xml");
    file = file.saveDlg("エクスポートファイル名を指定してください。",
			"*.xml",
			false);
    return file;
}



var analyzer = new Analyzer();
XML.prettyPrinting = false;
var xmldoc = new XML("<files/>");
var canceled = false;

if(app.documents.length > 0) {
    xmldoc.appendChild(analyzer.analyzeDoc(app.activeDocument));
}
else {
    var folderName = Folder.selectDialog("フォルダを選択してください。");
    if(folderName) {
	var folder = new Folder(folderName);
	var idFiles = selectIdFiles(folder);
	for(var i=0; i<idFiles.length; ++i) {
	    var doc = app.open(idFiles[i]);
	    xmldoc.appendChild(analyzer.analyzeDoc(doc));
	    doc.close(SaveOptions.NO);
	}
    }
    else {
	canceled = true;
    }
}

if(!canceled) {
    var file = getSaveFile();
    if(file && file.open("w")) {
	file.encoding = "utf-8";
	file.write("<?xml version='1.0' encoding='utf-8'?>");
	file.write(xmldoc.toXMLString());
	file.close();
    }
}

alert("finished");