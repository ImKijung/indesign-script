function functionCreateTOC() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '목차 생성 중: ';
    return function_core(CreateTOC, true, false, false);
}

function functionUpdateTOC() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '목차 업데이트 중: ';
    return function_core(UpdateTOC, true, false, false);
}
function functionUpdateTOCRtL() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '목차 업데이트 중: ';
    return function_core(UpdateTOCRtL, true, false, false);
}

function functionRemoveCharcterStyledText() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '스타일된 텍스트 삭제 중: ';
    return function_core(RemoveCharcterStyledText, true, false, false);
}


function CreateTOC(myDoc) {
    try {
        var tocStyles = myDoc.tocStyles;
        var nTocStyles = tocStyles.length;

        for (var i = 1; i < nTocStyles; i++)
            myDoc.createTOC(tocStyles[i], true);

        return RemoveCharcterStyledText(myDoc);
    } catch (ex) {
        errorMsgs.push('CreateTOC 함수 오류:(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function RemoveCharcterStyledText(myDoc, isRtL) {
    try {
        var result = true;
        if (!CheckTOCandRtL(myDoc, isRtL))
            return false;

        var noneStyle = myDoc.characterStyles[0];
        change_grep(myDoc, '\\t\\d{1,}\\-\\d{1,}\\r', null, null, null, null, noneStyle);

        if (isRtL == true) {
            var CLtR = myDoc.characterStyles.itemByName(TOCNumberStyles[0].attribute('name').toString());
            change_grep(myDoc, '(?<=\\t)\\d{1,}\\-\\d{1,}(?=\\r)', null, null, null, null, CLtR);
        }

        for (var i = 0 ; i < TOCremoveStyles.length() ; i++) {
            var charStyle = myDoc.characterStyles.itemByName(TOCremoveStyles[i].attribute('name').toString());
                if (charStyle != null && charStyle.isValid)
                    change_text(myDoc, null, null, charStyle, '', null, null);
        }

        return result;
    } catch (ex) {
        errorMsgs.push('RemoveCharcterStyledText 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function UpdateTOC(myDoc) {
    try {
        return UpdateTOCcommon(myDoc, false);
    } catch (ex) {
        errorMsgs.push('UpdateTOC 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function UpdateTOCRtL(myDoc) {
    try {
        return UpdateTOCcommon(myDoc, true);
    } catch (ex) {
        errorMsgs.push('UpdateTOCRtL 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function UpdateTOCcommon(myDoc, isRtL) {
    if (!CheckTOCandRtL(myDoc, isRtL))
        return false;
    var stories = myDoc.stories;
    for (var i = 0; i < stories.length; i++) {
        var story = stories[i]
        if(story.storyType != StoryTypes.TOC_STORY)
            continue;
        story.textContainers[0].select();
        app.menuActions.item("$ID/UpdateTableOfContentsCmd").invoke();
    }

    return RemoveCharcterStyledText(myDoc , isRtL);
}

function CheckTOCandRtL(myDoc, isRtL) {
    if (myDoc.name.search('TOC') < 0) {
        errorMsgs.push('문서가 TOC 문서가 아닙니다.');
        isStopped = true;
        return false;
    } else if (isRtL == true) {
        if (TOCNumberStyles.length() != 1) {
            errorMsgs.push('DB 설정 파일에 아라비아 문자 스타일 설정 값이 잘못되었습니다.');
            isStopped = true;
            return false;
        } else {
            var CLtR = myDoc.characterStyles.itemByName(TOCNumberStyles[0].attribute('name').toString());
            if (CLtR == null || !CLtR.isValid) {
                errorMsgs.push('TOC 문서에 ' + TOCNumberStyles[0].attribute('name').toString() + ' 문자 스타일이 없습니다.');
                isStopped = true;
                return false;
            }
        }
    }
    return true;
}