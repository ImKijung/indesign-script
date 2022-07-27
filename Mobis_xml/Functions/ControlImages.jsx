function functionResizelImages() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    if (!InputSizeData())
        return CloseWithoutMsg;
    progressTitle = '이미지 사이즈 조정 중: ';
    return function_core(ResizeImage, false, true, false);
}

function functionCheckImageSize() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '이미지 사이즈 확인 중: ';
    return function_core(CheckImageSize, false, false, false);
}

function functionAddImageSize() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'XML에 이미지 사이즈 입력 중: ';
    return function_core(AddImageSize, false, true, false);
}

function functionChangeImageFolder() {
    var docs = select_documents();
    if (!docs)
        return false;
    else if (!InputFolder('이미지 폴더 선택', docs[0]))
        return CloseWithoutMsg;
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '이미지 링크 재 설정 중: ';
    return function_core(ChangeImageLink, false, true, false);
}

function functionSetImageStyle() {
    SetImageStyle();
    return CloseWithoutMsg;
}

function functionAddImageStyle() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '이미지 문자 스타일 설정 중: ';
    return function_core(AddImageStyle, false, true, false);
}

function functionDelelteImageStyle() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '이미지 문자 스타일 삭제 중: ';
    return function_core(DeleteImageStyle, false, true, false);
}


function CheckImageSize(myDoc) {
    try {
        var allImages = myDoc.allGraphics;
        var imageCount = GetImageCountWithoutMaster(myDoc);
        for (var i = 0; i < imageCount; i++) {
            var image = allImages[i];
            var page = '알수 없는 페이지';
            if (image.parentPage != null)
                page = image.parentPage.name + 'page';
            if (image.absoluteVerticalScale != 100 || image.absoluteHorizontalScale != 100) {
                errorMsgs.push('이미지 사이즈가 100%가 아닙니다(' + myDoc.name + ':' + page + ', x:'
                 + image.absoluteHorizontalScale + ', y:' + image.absoluteVerticalScale + '): ' + image.itemLink.name);
                hasError = true;
            }
        }

        return true;
    } catch (ex) {
        errorMsgs.push('CheckImageSize 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function ResizeImage(myDoc) {
    try {
        var result = true;

        if (imageSize != 'UseAttribute' && typeof(imageSize) != 'number') {
            if (parseInt(imageSize) == NaN
                || parseInt(imageSize) == Infinity
                || parseInt(imageSize) == 0) {
                errorMsgs.push('이미지 사이즈가 올바르지 않습니다: ' + imageSize);
                isStopped = true;
                return false;
            }
        }

        var allImages = myDoc.allGraphics;
        var imageCount = GetImageCountWithoutMaster(myDoc);
        if (imageSize == 'UseAttribute') {
            for (var i = 0; i < imageCount; i++) {
                try {
                    try {
                        allImages[i].absoluteHorizontalScale = +allImages[i].associatedXMLElement.xmlAttributes.itemByName('xScale').value;
                        allImages[i].absoluteVerticalScale = +allImages[i].associatedXMLElement.xmlAttributes.itemByName('yScale').value;
                    } catch (ex) {
                        ex;
                    }
                    allImages[i].fit(FitOptions.CENTER_CONTENT);
                    allImages[i].fit(FitOptions.FRAME_TO_CONTENT);
                } catch (ex) {
                    ex;
                }
            }
        }
        else {
            for (var i = 0; i < imageCount; i++) {
                try {
                    allImages[i].absoluteVerticalScale = imageSize;
                    allImages[i].absoluteHorizontalScale = imageSize;
                    allImages[i].fit(FitOptions.CENTER_CONTENT);
                    allImages[i].fit(FitOptions.FRAME_TO_CONTENT);
                } catch (ex) { 
                    ex;
                }
            }
        }

        return result;
    } catch (ex) {
        errorMsgs.push('ResizeImage 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

// 문서의 XML구조에 이미지들의 크기를 특성값으로 넣어주는 함수
function AddImageSize(myDoc) {
    try {
        var result = true;

        var allImages = myDoc.allGraphics;
        var imageCount = GetImageCountWithoutMaster(myDoc);
        for (var i = 0; i < imageCount; i++) {
            try {
                allImages[i].associatedXMLElement.xmlAttributes.add('xScale', allImages[i].absoluteHorizontalScale.toString());
                allImages[i].associatedXMLElement.xmlAttributes.add('yScale', allImages[i].absoluteVerticalScale.toString());
                //이미지 오브젝트의 offset 값을 넣어주는 함수 by iM - 20.07.30
                var xoffset = allImages[i].parent.anchoredObjectSettings.anchorXoffset * 2.835;
                var yoffset = allImages[i].parent.anchoredObjectSettings.anchorYoffset * 2.835;
                allImages[i].associatedXMLElement.xmlAttributes.add('xOffset', xoffset + '');
                allImages[i].associatedXMLElement.xmlAttributes.add('yOffset', yoffset + '');
                allImages[i].associatedXMLElement.xmlAttributes.add('appliedObjectStyle', allImages[i].parent.appliedObjectStyle.name);
                //이미지 오브젝트의 offset 값을 넣어주는 함수 by iM - 20.07.30
            } catch (ex) {
                ex;
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('AddImageSize 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

// 이미지들의 링크 폴더를 실제로 수정해주는 함수
function ChangeImageLink(myDoc) {
    try {
        var result = true;

        //var allImages = myDoc.allGraphics;
        var imageCount = myDoc.allGraphics.length;
        for (var i = 0 ; i < imageCount ; i++) {
            try {
                linkName = myDoc.allGraphics[i].itemLink.name;
            } catch (ex) {
                errorMsgs.push('ChangeImageLink 함수 오류(getName): ' + myDoc.name + '문서 '
                 + myDoc.allGraphics[i].name + '이미지 - Line: ' + ex.line + ':: ' + ex);
                hasError = true;
                result = false;
                continue;
            }
            var imageFile = File(folderPath + '\\' + linkName);
            try {
                if (!imageFile.exists) {
                    if (folderPath2 == null) {
                        FixNonExist(myDoc);
                    }
                    if (folderPath2 != 'pass') {
                        imageFile = File(folderPath2 + '\\' + linkName);
                    }
                }
                myDoc.allGraphics[i].itemLink.relink(imageFile);
            } catch (ex) {
                errorMsgs.push('ChangeImageLink 함수 오류(reLink): ' + myDoc.name + '문서 '
                 + imageFile.fsName + '이미지 파일 - Line: ' + ex.line + ':: ' + ex);
                hasError = true;
                result = false;
            }
        }
        //allImages = myDoc.allGraphics;
        imageCount = myDoc.allGraphics.length;
        for (var i = 0 ; i < imageCount ; i++) {
            try {
                var imageFile = File(myDoc.allGraphics[i].itemLink.filePath);
                if (imageFile.exists)
                    myDoc.allGraphics[i].itemLink.update();
            } catch (ex) {
                errorMsgs.push('ChangeImageLink 함수 오류(update): ' + myDoc.name + '문서 '
                 + myDoc.allGraphics[i].name + '이미지 - Line: ' + ex.line + ':: ' + ex);
                hasError = true;
                result = false;
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('ChangeImageLink 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function FixNonExist(myDoc) {
    if (confirm ('선택한 폴더에 이미지가 없습니다. 추가 폴더를 설정하시겠습니까?', true, '추가 폴더 선택')) {
        var folder1 = folderPath;
        if (InputFolder('이미지 폴더 선택', myDoc)) {
            folderPath2 = folderPath;
            folderPath = folder1;
        } else {
            folderPath2 = 'pass';
        }
    } else
        folderPath2 = 'pass';
}

function GetImageCountWithoutMaster(myDoc) {
    try {
        var myMaster = myDoc.masterSpreads;
        var myMasterNumber = myMaster.count();
        var masterImages = 0;
        for (var i = 0 ; i < myMasterNumber ; i++) {
            imageNumber = myMaster[i].allGraphics;
            masterImages += imageNumber.length;
        }

        return myDoc.allGraphics.length - masterImages;
    } catch (ex) {
        errorMsgs.push('GetImageCountWithoutMaster 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function AddImageStyle(myDoc) {
    try {
        var result = true;
        
       //imageTagName
        var imageStyle = myDoc.characterStyles.itemByName(imageStyleName);
        if (imageStyle == null) {
            myDoc.importStyles(ImportFormat.CHARACTER_STYLES_FORMAT, inddFile);
            imageStyle = myDoc.characterStyles.itemByName(imageStyleName);
            if (imageStyle == null)
                throw new Error('ImageCharacterStyle.indd 파일에 ' + imageStyleName + ' 문자 스타일이 없습니다.');
        }
        
        var allImages = myDoc.allGraphics;
        var imageCount = GetImageCountWithoutMaster(myDoc) - 1;
        for ( ; imageCount >= 0 ; imageCount--) {
            try {
                var image = allImages[imageCount];
                var charObj = image.parent.parent;
                if (charObj.parent.toString() == '[object Story]')
                    charObj.applyCharacterStyle(imageStyle);
                else if (charObj.parent.toString() == '[object Cell]') {
                    var tableStyleName = charObj.parent.parent.appliedTableStyle.name;
                    for (var i = 0 ; i < styledImageTables.length() ; i++) {
                        if (tableStyleName == styledImageTables[i].attribute('name').toString() && CheckInline(charObj)) {
                            charObj.applyCharacterStyle(imageStyle);
                        }
                    }
                } else if (charObj.parent.toString() == '[object Document]') {
                    // pass
                    //
                } else if (charObj.toString() == '[object Group]') {
                    errorMsgs.push('그룹 설정을 적용한 이미지(' + myDoc.name + ':' + image.parentPage.name + 'page): ' + image.itemLink.name);
                    result = false;
                    hasError = true;
                } else {
                    alert(charObj.parent.toString() + '구조가 적용된 이미지를 확인(' + myDoc.name + ', ' + image.parentPage.name + 'page): ' + image.itemLink.name
                    + ', 정상구조는 [object Document], [object Story], [object Cell] 등의 구조를 가져야 합니다.', '스타일 및 구조 확인 필요.');
                }
            } catch (ex) {
                errorMsgs.push('AddImageStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
                hasError = true;
                result = false;
            }
        }

        return result;
    } catch (ex) {
        errorMsgs.push('AddImageStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function CheckInline(charObj) {
    return !(Trim(charObj.paragraphs[0].contents).length == 1); 
}

function DeleteImageStyle(myDoc) {
    try {
        var result = true;

        //imageTagName
        var imageStyle = myDoc.characterStyles.itemByName(imageStyleName);
        if (imageStyle != null) {
            var noneStyle = myDoc.characterStyles[0];
            imageStyle.remove(noneStyle);
        }

        return result;
    } catch (ex) {
        errorMsgs.push('DeleteImageStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function removeImageStyle(myDoc) {
    try {
        var result = true;

        //imageTagName
        var imageStyle = myDoc.characterStyles.itemByName(imageStyleName);
        if (imageStyle != null) {
            var noneStyle = myDoc.characterStyles[0];
            change_grep(myDoc, null, null, imageStyle, null, null, noneStyle)
        }

        return result;
    } catch (ex) {
        errorMsgs.push('DeleteImageStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetImageStyle() {
    try {
        app.open(inddFile.fsName);
    } catch (ex) {
        errorMsgs.push('SetImageStyle 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}