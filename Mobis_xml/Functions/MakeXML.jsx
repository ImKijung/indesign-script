function functionExportXML() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'XML 추출 중: ';
    var result = function_core(ExportXML, false, false, false);
    CloseAllInvisibleDocuments();
    return result;
}

function functionImportXML() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'XML 입력 중: ';
    var result = function_core(ImportXML, false, true, false);
    if (!hasError && !isStopped) {
        askSave = false;
        askBook = false;
        progressTitle = '상호참조 재 설정 중: ';
        result = function_core(SetHyperlinkFromXML, false, true, false);
    } else
        result = false;
    afterRemoveStructureAndTags();
    CloseAllInvisibleDocuments();
    return result;
}

function ExportXML(myDoc) {
    var layerVisibles = [];
    try {
        var result = true;
        if (myDoc.name.search('Cover') > 0 || myDoc.name.search('TOC') > 0)
            return result;

        SetLayerVisible(myDoc, layerVisibles, true);
        // var MMI = myDoc.conditions.item("MMI");
        // if (MMI.isValid) {
        //     MMI.visible = true;
        // } else { }
        if (!MakeFlatStructure (myDoc)) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에서 XML을 추출하지 못했습니다');
            hasError = true;
            return false;
        }

        CopyNotesToXML(myDoc);
        if (GetVisibleDocumentCount() == 1) {
            //open all indd file in same folder with invisible
            var files = new Folder(myDoc.filePath).getFiles('*.indd');
            for (var i = 0 ; i < files.length ; i++) {
                if (files[i].toString() != myDoc.fullName)
                    app.open(files[i], false);
            }
        } else if (app.books.length) {
            //open all document in book file with invisible
            var book = app.activeBook;
            var contents = book.bookContents;
            //set invisible contents codes....
            for (var i = 0 ; i < contents.length ; i++)
                app.open(contents[i].fullName, false);
        }

        if (!SaveHyperlinkToXML(myDoc)) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에서 XML을 추출하지 못했습니다');
            hasError = true;
            return false;
        }
        SaveIndexToXML(myDoc);

        myDoc.xmlExportPreferences.excludeDtd = false;
        myDoc.xmlExportPreferences.viewAfterExport = false;
        myDoc.xmlExportPreferences.exportFromSelected = false;
        myDoc.xmlExportPreferences.exportUntaggedTablesFormat = 1484022643; //CALS
        myDoc.xmlExportPreferences.characterReferences = false;
        myDoc.xmlExportPreferences.allowTransform = false;
        myDoc.xmlExportPreferences.fileEncoding = 1937134904; //UTF-8
        myDoc.xmlExportPreferences.ruby = false;

        var myName = null;
        var myFolder = null;
        try {
            myName = GetFileNameWithoutExtention(myDoc.name);
            myFolder = new Folder(myDoc.filePath + '/XML');
        } catch (ex) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name + ' 문서가 한 번도 저장되지 않았습니다');
            errorMsgs.push(myDoc.name + ' 문서에서 XML을 추출하지 못했습니다');
            hasError = true;
            return false;
        }
        try {
            myFolder.create();
        } catch (ex) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.filePath + '/XML 폴더 생성 오류: ' + ex);
            errorMsgs.push(myDoc.name + ' 문서에서 XML을 추출하지 못했습니다');
            hasError = true;
            return false;
        }

        var FileOutput = new File(myDoc.filePath + '/XML/' + myName + '.xml');
        if (indentXML.exists) {
            myDoc.xmlExportPreferences.allowTransform = true;
            myDoc.xmlExportPreferences.transformFilename = indentXML.fsName;
            myDoc.exportFile(ExportFormat.xml, FileOutput);
        } else {
            throw new Error('Reference 폴더에 indent_XML.xsl 파일이 없습니다. 프로그램이 손상되었습니다.');
        }

        /* var myXSLT;
        if (myXSLT) {
            myDoc.xmlExportPreferences.allowTransform = true;
            myDoc.xmlExportPreferences.transformFilename = myXSLT;
            FileOutput = new File(myDoc.filePath + '/XML/' + myName + '_MMI.xml');
            myDoc.exportFile(ExportFormat.xml, FileOutput);
            myDoc.xmlExportPreferences.allowTransform = false;
        } */
        SetLayerVisible(myDoc, layerVisibles, false);

        return result;
    } catch (ex) {
        errorMsgs.push('ExportXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        SetLayerVisible(myDoc, layerVisibles, false);
        throw ex;
    }
}

function ImportXML(myDoc) {
    var layerVisibles = [];
    try {
        var result = true;
        if (myDoc.name.search('Cover') > 0 || myDoc.name.search('TOC') > 0)
            return result;

        SetLayerVisible(myDoc, layerVisibles, true);
        if (GetVisibleDocumentCount() == 1) {
            //open all indd file in same folder with invisible
            var files = new Folder(myDoc.filePath).getFiles('*.indd');
            for (var i = 0 ; i < files.length ; i++) {
                if (files[i].toString() != myDoc.fullName)
                    app.open(files[i], false);
            }
            // if is working document => save destination to memory
            SaveDestinationToMemory(myDoc);
        } else if (app.books.length) {
            //open all document in book file with invisible
            var book = app.activeBook;
            var contents = book.bookContents;
            //set invisible contents codes....
            for (var i = 0 ; i < contents.length ; i++)
                app.open(contents[i].fullName, false);
        }

        if (!MakeFlatStructure (myDoc, true)) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
            hasError = true;
            return false;
        }
        
        myDoc.xmlImportPreferences.importStyle = 1481469289; //'XMmi'

        myDoc.xmlImportPreferences.createLinkToXML = false;
        myDoc.xmlImportPreferences.allowTransform = false;
        myDoc.xmlImportPreferences.repeatTextElements = false;
        myDoc.xmlImportPreferences.ignoreUnmatchedIncoming = false;
        myDoc.xmlImportPreferences.importTextIntoTables = false;
        myDoc.xmlImportPreferences.ignoreWhitespace = true;
        myDoc.xmlImportPreferences.removeUnmatchedExisting = true;
        myDoc.xmlImportPreferences.importCALSTables = true;
        
        myDoc.xmlImportPreferences.importToSelected = false;
        app.pdfPlacePreferences.pdfCrop = 1131573313; // art

        try {
            var myPath = myDoc.filePath;
        } catch (ex) {
            errorMsgs.push(myDoc.name + ' 문서를 한 번도 저장하지 않았습니다');
            errorMsgs.push(myDoc.name + ' 문서에 XML을 입력하지 못했습니다');
            SetLayerVisible(myDoc, layerVisibles, false);
            hasError = true;
            return false;
        }

        if (myDoc.xmlElements[0].xmlElements.count() < 1) {
            errorMsgs.push(myDoc.name + ' 문서가 XML구조를 가지고 있지 않습니다');
            errorMsgs.push(myDoc.name + ' 문서에 XML을 입력하지 못했습니다');
            SetLayerVisible(myDoc, layerVisibles, false);
            hasError = true;
            return false;
        }

        var xmlFile = new File(myPath + '/XML/' + GetFileNameWithoutExtention(myDoc.name) + '.xml');
        if (xmlFile.exists) {
            try {
                xmlFile.open('r');
                var xml = new XML(xmlFile.read());
                xmlFile.close();
                var hrefs = xml.xpath("//*[@href]");
                var hrefPass = true;
                for (var i = 0 ; i < hrefs.length() ; i++) {
                    var href = hrefs[i].attribute('href').toString();
                    var target = new File(href);
                    if (!target.exists) {
                        hrefPass = false;
                        errorMsgs.push(GetFileNameWithoutExtention(myDoc.name) + ' XML 문서 - 다음 파일이 없습니다: ' + target.fsName);
                    }
                }

                if (!hrefPass) {
                    hasError = true;
                    SetLayerVisible(myDoc, layerVisibles, false);
                    errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
                    return false;
                }
                // Delete conditional Text ID
                myDoc.conditions.everyItem().remove();
                DeleteHyperlink(myDoc);
                myDoc.importXML(xmlFile);
            } catch (ex) {
                var msg = ex.description;
                if (msg.search('href') >= 0)
                    msg = '손상된 링크가 있습니다. ' + msg;
                errorMsgs.push(GetFileNameWithoutExtention(myDoc.name) + '.xml 문서 오류: ' + msg);
                hasError = true;
                SetLayerVisible(myDoc, layerVisibles, false);
                errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
                return false;
            }
        } else {
            errorMsgs.push(GetFileNameWithoutExtention(myDoc.name) + ' 파일 이름과 똑같은 XML 파일이 없습니다');
            hasError = true;
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
            return false;
        }

        if (!ResizeImage(myDoc)) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
            hasError = true;
            return false;
        }

        //Import 일 경우 True 세팅. 원인 파악 중
        if (!MappingToStyle(myDoc, true)) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
            hasError = true;
            return false;
        }
        
        //chkSwatchName("ChangedMMI", myDoc);
        //AddChangedStyle(myDoc);
        //Changed_MMI(myDoc, "MMI");
        //Changed_MMI(myDoc, "MMI_NoBold");

        if (!SetPageBreak(myDoc)) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
            hasError = true;
            return false;
        }
        
        SetHyperlinkDestinationFromXML(myDoc);
        // if is working document => reconnect destination
        if (GetVisibleDocumentCount() == 1)
            ReconnectDestinationFromMemory(myDoc);

        if (!SetIndexFromXML(myDoc)) {
            SetLayerVisible(myDoc, layerVisibles, false);
            errorMsgs.push(myDoc.name.toString() + ' 문서에 XML을 입력하지 못했습니다');
            hasError = true;
            return false;
        }
        setMMIconditionID(myDoc);
        CopyNotesFromXML(myDoc);
        SetLayerVisible(myDoc, layerVisibles, false);
        return result;
    } catch (ex) {
        errorMsgs.push('ImportXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        CloseAllInvisibleDocuments();
        SetLayerVisible(myDoc, layerVisibles, false);
        throw ex;
    }
}
