function functionSetHighlight() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '하이라이트 적용 중: ';
    return function_core(SetHighlight, false, true, false);
}

function functionResetHighlight() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = '하이라이트 해제 중: ';
    return function_core(ResetHighlight, false, true, false);
}

function SetHighlight(myDoc) {
    try {
        var result = true;

        for (var i = 0 ; i < higthlightStyles.length() ; i++) {
            try {
                var styleType = higthlightStyles[i].localName();
                var styleName = higthlightStyles[i].attribute('name').toString();
                var style;
                var highlightColor = higthlightStyles[i].attribute('color').toString();
                var highlightedStyle;
                var highlightedStyleName = styleName + '_highlighted';
                if (styleType == 'CharacterStyle') {
                    style = myDoc.characterStyles.itemByName(styleName);
                    if (style == null) {
                        continue;
                    }
                    highlightedStyle = myDoc.characterStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle == null) {
                        highlightedStyle = myDoc.characterStyles.add({name: highlightedStyleName});
                        highlightedStyle.basedOn = style;
                    }
                    highlightedStyle.underline = true;
                    highlightedStyle.underlineTint = 100;
                    highlightedStyle.underlineWeight = THICK;
                    highlightedStyle.underlineOffset = '-3pt';
                    highlightedStyle.underlineColor = SetSwatch(myDoc, highlightedStyleName, highlightColor);

                    change_grep(myDoc, null, null, style, null, null, highlightedStyle);
                } else if (styleType == 'ParagraphStyle') {
                    style = myDoc.paragraphStyles.itemByName(styleName);
                    if (style == null) {
                        continue;
                    }
                    highlightedStyle = myDoc.paragraphStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle == null) {
                        highlightedStyle = myDoc.paragraphStyles.add({name: highlightedStyleName});
                        highlightedStyle.basedOn = style;
                    }
                    highlightedStyle.underline = true;
                    highlightedStyle.underlineTint = 100;
                    highlightedStyle.underlineWeight = THICK;
                    highlightedStyle.underlineOffset = '-3pt';
                    highlightedStyle.underlineColor = SetSwatch(myDoc, highlightedStyleName, highlightColor);

                    change_grep(myDoc, null, style, null, null, highlightedStyle, null);
                } else if (styleType == 'CellStyle') {
                    style = myDoc.cellStyles.itemByName(styleName);
                    if (style == null) {
                        continue;
                    }
                    highlightedStyle = myDoc.cellStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle == null) {
                        highlightedStyle = myDoc.cellStyles.add({name: highlightedStyleName});
                        highlightedStyle.basedOn = style;
                    }
                    highlightedStyle.fillTint = 100;
                    highlightedStyle.fillColor = SetSwatch(myDoc, highlightedStyleName, highlightColor);

                    ChangeCellStyle(myDoc, style, highlightedStyle);
                } else if (styleType == 'TableStyle') {
                    style = myDoc.tableStyles.itemByName(styleName);
                    if (style == null) {
                        continue;
                    }
                    highlightedStyle = myDoc.tableStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle == null) {
                        highlightedStyle = myDoc.tableStyles.add({name: highlightedStyleName});
                        highlightedStyle.basedOn = style;
                    }
                    highlightedStyle.fillTint = 100;
                    highlightedStyle.fillColor = SetSwatch(myDoc, highlightedStyleName, highlightColor);

                    ChangeTableStyle(myDoc, style, highlightedStyle);
                }
            } catch (ex) {
                errorMsgs.push('SetHighlight 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
                hasError = true;
                result = false;
            }
        }

        return result;
    } catch (ex) {
        errorMsgs.push('SetHighlight 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function ChangeCellStyle(myDoc, oldStyle, newStyle) {
    var stories = myDoc.stories;
    for (var i = 0 ; i < stories.length ; i++) {
        var story = stories[i];
        var tables = story.tables;
        ChangeCellStyles(tables, oldStyle, newStyle);
    }
}

function ChangeCellStyles(tables, oldStyle, newStyle) {
    if (tables != null) {
        for (var i = 0 ; i < tables.length ; i++) {
            var table = tables[i];
            var cells = table.cells;
            for (var j = 0 ; j < cells.length ; j++) {
                var cell = cells[j];
                if (cell.appliedTableStyle == oldStyle) {
                    cell.appliedTableStyle = newStyle;
                }
                ChangeTableStyles(cell.tables, oldStyle, newStyle);
            }
        }
    }
}

function ChangeTableStyle(myDoc, oldStyle, newStyle) {
    var stories = myDoc.stories;
    for (var i = 0 ; i < stories.length ; i++) {
        var story = stories[i];
        var tables = story.tables;
        ChangeTableStyles(tables, oldStyle, newStyle);
    }
}

function ChangeTableStyles(tables, oldStyle, newStyle) {
    if (tables != null) {
        for (var i = 0 ; i < tables.length ; i++) {
            var table = tables[i];
            if (table.appliedTableStyle == oldStyle) {
                table.appliedTableStyle = newStyle;
            }
            var cells = table.cells;
            for (var j = 0 ; j < cells.length ; j++) {
                var cell = cells[j];
                ChangeTableStyles(cell.tables, oldStyle, newStyle);
            }
        }
    }
}

function ResetHighlight(myDoc) {
    try {
        var result = true;

        for (var i = 0 ; i < higthlightStyles.length() ; i++) {
            try {
                var styleType = higthlightStyles[i].localName();
                var styleName = higthlightStyles[i].attribute('name').toString();
                var style;
                var highlightedStyle;
                var highlightedStyleName = styleName + '_highlighted';
                if (styleType == 'CharacterStyle') {
                    style = myDoc.characterStyles.itemByName(styleName);
                    highlightedStyle = myDoc.characterStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle != null) {
                        if (style == null)
                            throw new Error(styleName + ' 문자 스타일이 없습니다.');
                    }
                } else if (styleType == 'ParagraphStyle') {
                    style = myDoc.paragraphStyles.itemByName(styleName);
                    highlightedStyle = myDoc.paragraphStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle != null) {
                        if (style == null)
                            throw new Error(styleName + ' 단락 스타일이 없습니다.');
                    }
                } else if (styleType == 'CellStyle') {
                    style = myDoc.cellStyles.itemByName(styleName);
                    highlightedStyle = myDoc.cellStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle != null) {
                        if (style == null)
                            throw new Error(styleName + ' 셀 스타일이 없습니다.');
                    }
                } else if (styleType == 'TableStyle') {
                    style = myDoc.tableStyles.itemByName(styleName);
                    highlightedStyle = myDoc.tableStyles.itemByName(highlightedStyleName);
                    if (highlightedStyle != null) {
                        if (style == null)
                            throw new Error(styleName + ' 테이블 스타일이 없습니다.');
                    }
                } else
                    throw new Error('알 수 없는 스타일 종류: ' + styleType);
                if (highlightedStyle != null)
                    highlightedStyle.remove(style);
            } catch (ex) {
                errorMsgs.push('ResetHighlight 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
                hasError = true;
                result = false;
            }
        }

        return result;
    } catch (ex) {
        errorMsgs.push('ResetHighlight 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}


function SetSwatch(myDoc, swchName, swchColor) {
    try {
        //Create Swatch
        swchName = ReplaceAll(swchName, '-', '_');
        swchColor = ReplaceAll(swchColor, '#', '');
        var newSwatch = myDoc.swatches.itemByName(swchName);
        if (newSwatch == null) {
            newSwatch = myColorAdd(myDoc, swchName, ColorModel.PROCESS, swchColor);
        }

        return newSwatch;
    } catch (ex) {
        errorMsgs.push('SetSwatch 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function myColorAdd(myDoc, myColorName, myColorModel, myColorValue){
    if (myColorValue instanceof Array == false) {
        myColorValue = [(parseInt(myColorValue, 16) >> 16) & 0xff, (parseInt(myColorValue, 16) >> 8) & 0xff, parseInt(myColorValue, 16) & 0xff ];
        myColorSpace = ColorSpace.RGB;
    } else {
        if (myColorValue.length == 3)
            myColorSpace = ColorSpace.RGB;
        else
            myColorSpace = ColorSpace.CMYK;
    }

    try {
        myColor = myDoc.colors.item(myColorName);
        myName = myColor.name;
    } catch (myError) {
        myColor = myDoc.colors.add();
        myColor.properties = {name:myColorName, model:myColorModel, space:myColorSpace ,colorValue:myColorValue};
    }

    return myColor;
}
