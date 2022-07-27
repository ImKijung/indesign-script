function functionCheckAndMakeTags() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'XML 구조 및 태그 생성 중: ';
    return function_core(CheckAndMakeTags, false, true, false);
}

function CheckAndMakeTags(myDoc) {
    try {
        var result = true;

        if (!RemoveStructureAndTags(myDoc)) {
            hasError = true;
            return false;
        }

        //단락 스타일의 숫자만큼 체크. 배열숫자 때문에 1을 빼줌
        if (!CheckErrors(myDoc)) {
            hasError = true;
            return false;
        }

        //0번 => 단락 스타일 없슴.
        //1번 => 기본 단락 스타일.
        var allParagraphStyles = myDoc.allParagraphStyles;
        for (var i = 2 ; i < allParagraphStyles.length ; i++) {
            var myPaStyle = allParagraphStyles[i];        // 단락 스타일
            result = result && AddToTags(myDoc, myPaStyle); // 태그와 단락을 비교, 추가
        }
        if (!result) {
            hasError = true;
            return false;
        }

        //0번 => 문자 스타일 없슴.
        var allCharacterStyles = myDoc.allCharacterStyles;
        for (var i = 1 ; i < allCharacterStyles.length ; i++) {
            var myChStyle = allCharacterStyles[i];
            result = result && AddToTags(myDoc, myChStyle);
        }

        //0번 => 표 스타일 없슴.
        //1번 => 기본 표 스타일.
        var allTableStyles = myDoc.allTableStyles;
        for (var i = 2 ; i < allTableStyles.length ; i++) {
            var myStyle = allTableStyles[i];
            result = result && AddToTags(myDoc, myStyle);
        }

        //0번 => 셀 스타일 없슴.
        var allCellStyles = myDoc.allCellStyles;
        for (var i = 1 ; i < allCellStyles.length ; i++) {
            var myStyle = allCellStyles[i];
            result = result && AddToTags(myDoc, myStyle);
        }
        
        if (!result) {
            hasError = true;
            return false;
        }

        return result;
    } catch (ex) {
        errorMsgs.push('CheckAndMakeTags 함수 오류(' + myDoc.name + '): ' + ex);
        hasError = true;
        throw ex;
    }
}

// 스타일을 태그에 추가해주는 함수
function AddToTags(myDoc, myStyle) {
    try {
        var result = true;

        var myTag = null;
        var myTagName = myStyle.name;
        // Make a Tag with the Style name..
        try {
            myTag = myDoc.xmlTags.itemByName(myTagName);
            if (myTag == null)
                myTag = myDoc.xmlTags.add({name: myTagName});
        } catch (ex) {    // Tag가 있는 경우 충돌을 핸들링
            errorMsgs.push(myDoc.name + ' 의 ' + myStyle.name + '은(는) 태그화 할 수 있는 정상적인 스타일 이름이 아닙니다: ' + ex);
            hasError = true;
            return false;
        }
        myDoc.xmlExportMaps.add(myStyle, myTag);

        return result;
    } catch (ex) {
        errorMsgs.push('AddToTags 함수 오류(' + myDoc.name + '): ' + ex);
        hasError = true;
        throw ex;
    }
}