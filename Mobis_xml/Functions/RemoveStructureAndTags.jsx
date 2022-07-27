function functionRemoveStructureAndTags() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'XML 구조 및 태그 제거 중: ';
    return function_core(RemoveStructureAndTags, false, true, false);
}

function RemoveStructureAndTags(myDoc) {
    return StructureRemove(myDoc) && DeleteTags(myDoc);
}

function afterRemoveStructureAndTags() {
    try {
        var book = app.activeBook;
        var contents = book.bookContents;
        for (var n=0; n<contents.length; n++) {
            var myDoc = app.open(contents[n].fullName, false);
            StructureRemove(myDoc);
            DeleteTags(myDoc);
            myDoc.close(SaveOptions.YES);
        }
        return true;
    } catch (ex) {
        errorMsgs.push('afterRemoveStructureAndTags 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function DeleteUnusedTags(myDoc) {
    try {
        myDoc.deleteUnusedTags();
        return true;
    } catch (ex) {
        errorMsgs.push('DeleteUnusedTags 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function DeleteTags(myDoc) {
    try {
        if (myDoc.xmlTags.length > 1) {
            var start = 0;
            var count = myDoc.xmlTags.length;
            for ( ; count > 0 ; count --) {
                if (myDoc.xmlTags[start].name != 'Root') {
                    myDoc.xmlTags[start].remove('Root');
                } else {
                    start++;
                }
            }
        }
        return true;
    } catch (ex) {
        errorMsgs.push('DeleteTags 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function StructureRemove(myDoc) {
    try  {
        var root = myDoc.xmlElements[0];
        
        if (root.xmlElements.length > 0) {
            var count = root.xmlElements.length;
            for ( ; count > 0 ; count--)
                root.xmlElements[0].untag();
        }
        if (root.xmlAttributes.length > 0) {
            var count = root.xmlAttributes.length;
            for ( ; count > 0 ; count--)
                root.xmlAttributes[0].remove();
        }
        return true;
    }catch (ex) {
        errorMsgs.push('StructureRemove 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}