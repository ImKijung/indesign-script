function functionExportIDML() {
    var docs = select_documents();
    if (!docs)
        return false;
    else if (!SaveFolder('IDML 저장 폴더 선택', docs[0]))
        return false;
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'IDML 파일 내보내는 중: ';
    return function_core(ExportIDML, false, true, false);
}

function ExportIDML(myDoc) {
    try {
        var result = true;
        //if (myDoc.name.search('Cover') > 0 || myDoc.name.search('TOC') > 0)
            //return result;

        var idmlFile = new File(folderPath + '\\' + GetFileNameWithoutExtention(myDoc.name) + '.idml');
        myDoc.exportFile(ExportFormat.INDESIGN_MARKUP, idmlFile, true);

        return result;
    } catch (ex) {
        errorMsgs.push('ExportIDML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}