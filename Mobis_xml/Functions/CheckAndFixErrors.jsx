function functionCheckErrors() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '문서 오류 확인 중: ';
    return function_core(CheckErrors, false, true, false);
}

function functionFixErrors() {
    progressTitle = '오류 수정 중(double space, styled enter): ';
    return function_core(FixErrors, false, true, false);
}

function CheckErrors(myDoc) {
    try {
        var result = true;
        
        result = result && FindIncorrectParagraphStyle(myDoc);
        result = result && FindIncorrectCharacterStyle(myDoc);
        result = result && FindIncorrectTableStyle(myDoc);
        result = result && FindIncorrectCellStyle(myDoc);
        //result = result && FindIncorrectEnterCharacterStyle(myDoc);
        result = result && FindIncorrectImageCharacterStyle(myDoc);

        return result;
    } catch (ex) {
        errorMsgs.push('CheckErrors 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function FixErrors(myDoc, fix) {
    try {
        var result = true;
        
        //fix double space
        //FixDoubleSpace(myDoc, fix);
        //fix styled enter
        FixStyledEnter(myDoc, fix);

        return result;
    } catch (ex) {
        errorMsgs.push('FixErrors 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

// 문서의 단락 스타일에 [없음]과 [기본단락]이 사용되었는지 찾는 함수
function FindIncorrectParagraphStyle(myDoc) {
    try {
        var result = true;

        var allParagraphStyles = myDoc.allParagraphStyles;
        app.findTextPreferences = NothingEnum.NOTHING;
        app.findTextPreferences.appliedParagraphStyle = allParagraphStyles[1];
        
        var detect = myDoc.findText();
        var detect_Length = detect.length;
        if (detect_Length > 0) {
            detect_Length = 0;
            for (var i = 0 ; i > detect.length ; i++) {
                if (detect[i].parent.toString() != '[object Note]') {
                    detect_Length = detect_Length + 1;
                }
            }
            if (detect_Length > 0) {
                errorMsgs.push(myDoc.name + ' 문서에 ' + allParagraphStyles[1].name + '이(가) 사용되었습니다.');
                result = false;
                hasError = true;
            }
        }
        app.findTextPreferences.appliedParagraphStyle = allParagraphStyles[0];
        detect = myDoc.findText();
        detect_Length = detect.length;
        if (detect_Length > 0) {
            detect_Length = 0;
            for (var i = 0 ; i > detect.length ; i++) {
                if (detect[i].parent.toString() != '[object Note]') {
                    detect_Length = detect_Length + 1;
                }
            }
            if (detect_Length > 0) {
                errorMsgs.push(myDoc.name + ' 문서에 ' + allParagraphStyles[0].name + '이(가) 사용되었습니다.');
                result = false;
                hasError = true;
            }
        }
        app.findTextPreferences = NothingEnum.NOTHING;

        for (var i = 2 ; i < allParagraphStyles.length ; i++) {
            if (allParagraphStyles[i].name != allParagraphStyles[i].name.replace(/\s+/g, '')) {
                errorMsgs.push(myDoc.name + ' 문서의 ' + allParagraphStyles[i].name + ' 단락 스타일에 공백이 사용되었습니다.');
                result = false;
                hasError = true;
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('FindIncorrectParagraphStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        app.findTextPreferences = NothingEnum.NOTHING;
        throw ex;
    }
}

function FindIncorrectCharacterStyle(myDoc) {
    try {
        var result = true;

        var allCharacterStyles = myDoc.allCharacterStyles;
        for (var i = 1 ; i < allCharacterStyles.length ; i++) {
            if (allCharacterStyles[i].name != allCharacterStyles[i].name.replace(/\s+/g, '')) {
                errorMsgs.push(myDoc.name + ' 문서의 ' + allCharacterStyles[i].name + ' 문자 스타일에 공백이 사용되었습니다.');
                result = false;
                hasError = true;
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('FindIncorrectCharacterStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function FindIncorrectTableStyle(myDoc) {
    try {
        var result = true;

        var allTableStyles = myDoc.allTableStyles;
        for (var i = 2 ; i < allTableStyles.length ; i++) {
            if (allTableStyles[i].name != allTableStyles[i].name.replace(/\s+/g, '')) {
                errorMsgs.push(myDoc.name + ' 문서의 ' + allTableStyles[i].name + ' 표 스타일에 공백이 사용되었습니다.');
                result = false;
                hasError = true;
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('FindIncorrectTableStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function FindIncorrectCellStyle(myDoc) {
    try {
        var result = true;

        var allCellStyles = myDoc.allCellStyles;
        for (var i = 2 ; i < allCellStyles.length ; i++) {
            if (allCellStyles[i].name != allCellStyles[i].name.replace(/\s+/g, '')) {
                errorMsgs.push(myDoc.name + ' 문서의 ' + allCellStyles[i].name + ' 셀 스타일에 공백이 사용되었습니다.');
                result = false;
                hasError = true;
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('FindIncorrectCellStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}


function FindIncorrectEnterCharacterStyle(myDoc) {
    try {
        var result = true;
        if (myDoc.name.search('Cover') > 0 || myDoc.name.search('TOC') > 0 || myDoc.name.search('Index') > 0)
            return true;

        var allCharacterStyles = myDoc.allCharacterStyles;
        app.findTextPreferences = NothingEnum.NOTHING;
        app.findTextPreferences.findWhat = '^b';
        for (var i = 1 ; i < allCharacterStyles.length ; i++) {
            app.findTextPreferences.appliedCharacterStyle = allCharacterStyles[i];
            var detect = myDoc.findText();
            var detect_Length = detect.length;
            if (detect_Length) {
                errorMsgs.push(myDoc.name + ' 문서에 ' + allCharacterStyles[i].name + '이(가) 적용된 엔터가 ' + detect_Length + '개 있습니다.');
                result = false;
                hasError = true;
            }
        }

        /*
        app.findTextPreferences.appliedCharacterStyle = NothingEnum.NOTHING;
        app.findTextPreferences.findWhat = '^n';
        detect = myDoc.findText();
        detect_Length = detect.length;
        if (detect_Length) {
            errorMsgs.push(myDoc.name + ' 문서에 강제 줄넘김 엔터가 ' + detect_Length + '개 있습니다.');
            result = false;
        }
        */
        app.findTextPreferences = NothingEnum.NOTHING;

        return result;
    } catch (ex) {
        errorMsgs.push('FindIncorrectEnterCharacterStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        app.findTextPreferences = NothingEnum.NOTHING;
        throw ex;
    }
}

function FindIncorrectImageCharacterStyle(myDoc) {
    try {
        var result = true;

        var allGraphics = myDoc.allGraphics;
        var myMaster = myDoc.masterSpreads;
        var myMasterNumber = myMaster.count() - 1;
        var masterImage = 1;
        for ( ; myMasterNumber > -1 ; myMasterNumber--) {
            imageNumber = myMaster[myMasterNumber].allGraphics;
            masterImage += imageNumber.length;
        }
        var imageCount = allGraphics.length - masterImage;

        var styleNames = [];
        for (var i = 0 ; i < styledImageTables.length() ; i++) {
            styleNames[i] = styledImageTables[i].attribute('name').toString();
        }
        var styleName = styleNames.join('/');

        if (imageCount >= 0 && checkImageStyled) {
            var imageStyle = myDoc.characterStyles.itemByName(imageStyleName);
            for ( ; imageCount > -1 ; imageCount--) {
                var image = allGraphics[imageCount];
                var charObj = image.parent.parent;
                if (charObj.parent.toString() == '[object Story]') {
                    if (charObj.appliedCharacterStyle.name != imageStyleName) {
                        errorMsgs.push(imageStyleName + '을(를) 적용하지 않은 이미지(' + myDoc.name + ':' + image.parentPage.name + 'page): ' + image.itemLink.name);
                        result = false;
                        hasError = true;
                    }
                } else if (charObj.parent.toString() == '[object Cell]') {
                    var cellStyleName = charObj.parent.appliedCellStyle.name;
                    var tableStyleName = charObj.parent.parent.appliedTableStyle.name;
                    var matched = false;
                    for (var i = 0 ; i < styledImageTables.length() ; i++) {
                        if (tableStyleName == styledImageTables[i].attribute('name').toString() && CheckInline(charObj)) {
                            matched = true;
                            break;
                        }
                    }
                    if (matched) {
                        if (charObj.appliedCharacterStyle.name != imageStyleName) {
                            errorMsgs.push(imageStyleName + '을(를) 적용하지 않은 이미지(' + myDoc.name + ':' + image.parentPage.name + 'page): ' + image.itemLink.name
                             + '(적용한 표 스타일:' + tableStyleName + ')');
                            result = false;
                            hasError = true;
                        }
                    } else if (charObj.appliedCharacterStyle.name == imageStyleName) {
                        var emsg = imageStyleName + '을(를) 잘못 적용한 이미지(' + myDoc.name + ':' + image.parentPage.name + 'page): ' + image.itemLink.name + '(';
                        if (CheckInline(charObj))
                            emsg = emsg + '인라인 이미지, ';
                        emsg = emsg + '적용한 표 스타일:' + tableStyleName + ', 등록된 표 스타일:' + styleName + ')';
                        errorMsgs.push(emsg);
                        result = false;
                        hasError = true;
                    }
                } else if (charObj.parent.toString() == '[object Document]') {
                    //
                    //
                } else if (charObj.toString() == '[object Group]') {
                    errorMsgs.push('그룹 설정을 적용한 이미지(' + myDoc.name + ':' + image.parentPage.name + 'page): ' + image.itemLink.name);
                    result = false;
                    hasError = true;
                } else {
                    alert(charObj.parent.toString() + '구조가 적용된 이미지를 확인(' + myDoc.name + ':' + image.parentPage.name + 'page): ' + image.itemLink.name
                    + ', 정상구조는 [object Document], [object Story], [object Cell] 등의 구조를 가져야 합니다.', '스타일 및 구조 확인 필요.');
                }
            }
        }

        return result;
    } catch (ex) {
        errorMsgs.push('FindIncorrectImageCharacterStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function FixDoubleSpace(myDoc, fix) {
    try {
        var result = true;
        if (myDoc.name.search('Cover') > 0 || myDoc.name.search('TOC') > 0
         || myDoc.name.search('Index') > 0 || myDoc.name.search('Readmefirst') > 0)
            return result;

        var grep = '[~m~>~f~|~S~s~<~/~.~3~4~% ]{2,}';
        var changed = change_grep(myDoc, grep, null, null, '\\s', null, null);
        if (changed != null && changed.length != 0)
            errorMsgs.push(myDoc.name + ' 문서에서 ' + changed.length + '개의 Double Space 오류를 수정했습니다.');

        return result;
    } catch (ex) {
        errorMsgs.push('FixDoubleSpace 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function FixStyledEnter(myDoc, fix) {
    try {
        var result = true;
        if (myDoc.name.search('Cover') > 0 || myDoc.name.search('TOC') > 0 || myDoc.name.search('Index') > 0)
            return result;

        var noneStyle = myDoc.allCharacterStyles[0];
        var totalLength = 0;
        for (var i = 1 ; i < myDoc.characterStyles.length ; i++) {
            var charStyle = myDoc.characterStyles[i];
            var changed = change_grep(myDoc, '\\r', null, charStyle, null, null, noneStyle);
            if (changed != null)
                totalLength += changed.length;
        }
        if (fix == true && totalLength != 0)
            errorMsgs.push(myDoc.name + ' 문서에서 ' + totalLength + '개의 Styled Enter 오류를 수정했습니다.');
        
        return result;
    } catch (ex) {
        errorMsgs.push('FixStyledEnter 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}