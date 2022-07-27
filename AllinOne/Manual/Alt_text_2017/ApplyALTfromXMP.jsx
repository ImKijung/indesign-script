/***********************************************************************************************************
        Apply Alt Tag to Document Images
        Version : 1.1
        Type : Script
        InDesign : CS5.5
        Idea : Anne-Marie Concepcion | http://indesignsecrets.com
        Author : Marijan Tompa (tomaxxi) | Subotica [Serbia]
        Date : 12/07/2011
        Contact : me (at) tomaxxi (dot) com
        Twitter: @tomaxxi
        Web : http://tomaxxi.com
***********************************************************************************************************/

var appName = "Object ALT Text Changer | by @tomaxxi | edit by @geonwoo";

if(parseFloat(app.version) > 7){
    if(app.documents.length < 1){
        alert("No Documents Opened!", appName);
        exit();
    }else{
        var doc = app.activeDocument;
        if(doc.links.length < 1){
            alert("No Links found!", appName);
            exit();
        }
    }
    addAltText();
}else{
    alert("Adobe InDesign CS5.5 or later version is supported!", appName);
    exit();
}

function getXMLfromExcel() {
    
    var cScript = new File($.fileName);
    var path = cScript.parent.fsName;
    
    var script = new File(path + '\\GetXmlStringFromExcel.vbs');
    return XML(app.doScript (script, ScriptLanguage.VISUAL_BASIC, [path + '\\Manual_Metadata.xlsx'])); 
}

function addAltText(){
    
    var xml = getXMLfromExcel();
    var langs = [];
    var cs = xml.xpath ('//Row[1]').children();
    for (var i=0 ; i<cs.length() ; i++) {
        langs.push(cs[i].name());
    }
    
    var links = doc.links;
    
    var sources = ["From Structure", "XMP : Title", "XMP : Description", "XMP : Headline"];
    var values = [SourceType.SOURCE_XML_STRUCTURE, SourceType.SOURCE_XMP_TITLE, SourceType.SOURCE_XMP_DESCRIPTION, SourceType.SOURCE_XMP_HEADLINE];
    var dlg = new Window('dialog', appName);
    var g0 = dlg.add('group');
    var p0 = g0.add('panel');
        p0.orientation = 'row';
        p0.add('statictext', undefined, "Alt Text Source:");
        var lang = p0.add('dropdownlist', undefined, langs);
            lang.selection = 0;
        //var source = p0.add('dropdownlist', undefined, sources);
            //source.selection = 2;

    var g1 = dlg.add('group');
    var ok = g1.add('button', undefined, "OK");
    var canc = g1.add('button', undefined, "Cancel");
    var res = dlg.show();
    
    if(res == true){
        var l = lang.selection;
        var i, counter = 0;
        var link = null;
        var flieName = null;
        try {
            for(i = 0; i < links.length; i++){
                link = links[i];
                //link.parent.parent.objectExportOptions.altTextSourceType = values[source.selection.index];
                
                //Set alt from image
                //link.parent.parent.objectExportOptions.altMetadataProperty = ['dc', 'dc:description[@xml:lang="' + l.toString() + '"]'];
                //link.parent.parent.objectExportOptions.altTextSourceType = SourceType.SOURCE_XMP_OTHER;
                
                //Set alt from XML
                fileName = getFilenameFromPath(link.filePath);
                link.parent.parent.objectExportOptions.altTextSourceType = SourceType.SOURCE_CUSTOM;
                var item = xml.xpath('//Row[@imgName="' + fileName + '"]/' + l.toString() + '/text()').toString();
                link.parent.parent.objectExportOptions.customAltText = item;
            /*
                if (item != '')
                    link.parent.parent.objectExportOptions.customAltText = item;
                else {
                    //throw Error('엑셀 파일에 ' + fileName + ' 이미지의 ' + l.toString() + ' 언어 정보가 없습니다. 파일을 확인해 주십시오');
                    alert('엑셀 파일에 ' + fileName + ' 이미지의 ' + l.toString() + ' 언어 정보가 없습니다. 파일을 확인해 주십시오');
                    return false;
                }
                */
                //alert( link.parent.parent.objectExportOptions.altText());
                counter++;
            }
            if(counter > 0)
                alert("Complete: " + lang.selection.text + ", " + counter, appName);
            else
                alert("No Images affected!", appName);
        } catch (ex) {
            alert(ex.toString());
        }
    }
}

function getFilenameFromPath(fullPath) {
    return fullPath.replace(/^.*[\\\/]/, '');
}    
