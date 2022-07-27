function main_stylename() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    return function_core(stylename_trim, false, true, false);
}

function stylename_trim(myDoc) {
    try {
        var paras = myDoc.allParagraphStyles;
        for (var i = 2 ; i < paras.length ; i++) {
            var para = paras[i];
            para.name = para.name.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
        }

        var charas = myDoc.allCharacterStyles;
        for (var i = 1 ; i < charas.length ; i++) {
            var chara = charas[i];
            chara.name = chara.name.replace(/^[\s\uFEFF\xA0]+|[\s\uFEFF\xA0]+$/g, '');
        }

        return true;
    } catch (ex) {
        errorMsgs.push('stylename_trim 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}