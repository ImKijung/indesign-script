function functionMakeFlatStructure() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'XML 구조 생성 중: ';
    return function_core(MakeFlatStructure, false, true, false);
}

//XML구조생성
function MakeFlatStructure(myDoc, fix) {
    try {
        var result = true;

        myDoc.xmlPreferences.defaultCellTagName = "Cell";
        myDoc.xmlPreferences.defaultImageTagName = "Image";
        myDoc.xmlPreferences.defaultStoryTagName = "Story";
        myDoc.xmlPreferences.defaultTableTagName = "Table";
        
        var myGraphics = myDoc.allGraphics;
        var myGraphicsCount = myGraphics.length;
        var myMaster = myDoc.masterSpreads;
        var myMasterNumber = myMaster.count() - 1;
        var masterImage = 0;
        for ( ; myMasterNumber > -1 ; myMasterNumber--) {
            imageNumber = myMaster[myMasterNumber].allGraphics;
            masterImage += imageNumber.length;
        }
        myGraphicsCount -= masterImage;

        for (var j = 0; j < myGraphicsCount ; j++) {
            try {
                if (myGraphics[j].itemLink.status == 1819109747) {
                    // 링크 오류 상태
                    errorMsgs.push('MakeFlatStructure 함수 오류: ' + myDoc.name + '문서에 이미지링크 오류가 있습니다. 링크를 재설정 해주세요');
                    hasError = true;
                    return false;
                }
            } catch (ex) {
                throw ex;
            }
        }
        if (!FixErrors(myDoc, fix)) {
            hasError = true;
            return false;
        }
        if (!CheckAndMakeTags(myDoc)) {
            hasError = true;
            return false;
        }
            
        for (var i = 0 ; i < myGraphicsCount ; i++) {
            try {
                myGraphics[i].autoTag();
            } catch (ex) {
                ex;
                continue;
            }
        }

        //MappingToStyle에 넣지 말것
        myDoc.mapStylesToXMLTags();  // 구조 등록
        myDoc.mapStylesToXMLTags();  // 이미지를 위해 한번 더 구조 등록
        //MappingToStyle에 넣을 경우 중복 호출로 인해 xml 구조 소실
        
        //MappingToStyle에 Table 속성이 사라지는 문제로 순서 변경 by iM 20.08.03
        if (!MappingToStyle(myDoc)){
            hasError = true;
            return false;
        }
        
        SetMMIattribute(myDoc);
        
        if (!AddImageSize(myDoc)) {
            hasError = true;
            return false;
        }
        if (!SetPageBreakAttr(myDoc)) {
            hasError = true;
            return false;
        }
        
        return result;
    } catch (ex) {
        errorMsgs.push('MakeFlatStructure 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function MappingToStyle(myDoc, imporCheck) {
    try {
        var result = true;
        
        // var tables = myDoc.xmlElements[0].evaluateXPathExpression("descendant::Table[not(@HeaderRowCount)]");
        var tables = myDoc.xmlElements[0].evaluateXPathExpression(".//*[contains(name(),'Table_')]"); // 요소의 이름만으로 테이블 요소를 찾도록 수정 by iM 20.08.03
        var tables_count = tables.length;
        if (tables_count > 0) {
            for (var i=0 ; i<tables_count ; i++) {
                var table = tables[i].tables[0];
                var tbDirection = ""; //추가 속성 by iM - 20.07.30
                var rows = table.rows;
                var rows_count = rows.length;
                for (var j=0 ; j<rows_count ; j++)
                    rows[j].cells[0].associatedXMLElement.xmlAttributes.add('newrow', 'newrow');
                var columns = table.columns;
                var columns_count = columns.length;
                var widths = [];
                for (var j=0 ; j<columns_count ; j++)
                    widths = widths.concat((columns[j].width/table.width).toFixed(2) + '*');
                tables[i].xmlAttributes.add('colspecs', widths.join(':'));
                tables[i].xmlAttributes.add('HeaderRowCount', table.headerRowCount + '');
                //추가 속성 by iM - 20.07.30
                tables[i].xmlAttributes.add('FooterRowCount', table.footerRowCount + '');
                tables[i].xmlAttributes.add('BodyRowCount', table.bodyRowCount + '');
                tables[i].xmlAttributes.add('ColumnCount', table.columnCount + '');
                if (table.tableDirection == "1278366308") {
                    tbDirection = "LeftToRightDirection";
                } else
                    tbDirection = "RightToLeftDirection";
                tables[i].xmlAttributes.add('TableDirection', tbDirection);
                tables[i].xmlAttributes.add('AppliedTableStyle', table.appliedTableStyle.name);
                //추가 속성 by iM - 20.07.30
            }

            // var cells = myDoc.xmlElements[0].evaluateXPathExpression("descendant::Cell[not(@namest)]");
            var cells = myDoc.xmlElements[0].evaluateXPathExpression(".//*[contains(name(),'Cell')]"); // 요소의 이름만으로 셀 요소를 찾도록 수정 by iM 20.08.03
            var cells_count = cells.length;
            for (var i=0 ; i<cells_count ; i++) {
                var cell = cells[i].cells[0];
                var cellfBaseline = ""; //추가 속성 by iM - 20.07.30
                var cellVerticajust = ""; //추가 속성 by iM - 20.07.30
                //추가 속성 by iM - 20.07.30
                cells[i].xmlAttributes.add('Name', cell.name + '');
                cells[i].xmlAttributes.add('RowSpan', cell.rowSpan + '');
                cells[i].xmlAttributes.add('ColumnSpan', cell.columnSpan + '');
                cells[i].xmlAttributes.add('AppliedCellStyle', cell.appliedCellStyle.name);
                if (cell.firstBaselineOffset == "1296135023") {
                    cellfBaseline = "ASCENT_OFFSET";
                } if (cell.firstBaselineOffset == "1296255087") {
                    cellfBaseline = "CAP_HEIGHT";
                } if (cell.firstBaselineOffset == "1296386159") {
                    cellfBaseline = "EMBOX_HEIGHT";
                } if (cell.firstBaselineOffset == "1313228911") {
                    cellfBaseline = "FIXED_HEIGHT";
                } if (cell.firstBaselineOffset == "1296852079") {
                    cellfBaseline = "LEADING_OFFSET";
                } if (cell.firstBaselineOffset == "1299728495") {
                    cellfBaseline = "X_HEIGHT";
                }
                cells[i].xmlAttributes.add('FirstBaselineOffset', cellfBaseline);
                if (cell.verticalJustification == "1953460256") {
                    cellVerticajust = "TOP_ALIGN";
                } if (cell.verticalJustification == "1667591796") {
                    cellVerticajust = "CENTER_ALIGN";
                } if (cell.verticalJustification == "1651471469") {
                    cellVerticajust = "BOTTOM_ALIGN";
                } if (cell.verticalJustification == "1785951334") {
                    cellVerticajust = "JUSTIFY_ALIGN";
                }
                cells[i].xmlAttributes.add('VerticalJustification', cellVerticajust);
                //추가 속성 by iM - 20.07.30
                if (cell.columnSpan > 1) {
                    var namest = cell.parentColumn.index + 1;
                    var nameend = namest + cell.columnSpan - 1;
                    cells[i].xmlAttributes.add('namest', 'col' + namest);
                    cells[i].xmlAttributes.add('nameend', 'col' + nameend);
                }
            }
        }

        var TagList = myDoc.xmlTags.everyItem();
        var TagCounter = myDoc.xmlTags.count();
        var None = myDoc.characterStyles[0];
        var allParagraphStyles = myDoc.allParagraphStyles;
        var allCharacterStyles = myDoc.allCharacterStyles;
        var allTableStyles = myDoc.allTableStyles;
        var allCellStyles = myDoc.allCellStyles;
        var xmlMaps = myDoc.xmlImportMaps;

        for (var i = 0 ; i < TagCounter ; i++) {
            // Cell, Story, Table, Topic*, PA_* 들은 따로 처리 해준다.
            TagName = TagList.name[i];
            var added = false;
            
            if (TagName == 'Root'
            || TagName == myDoc.xmlPreferences.defaultCellTagName
            || TagName == myDoc.xmlPreferences.defaultImageTagName
            || TagName == myDoc.xmlPreferences.defaultStoryTagName
            || TagName == myDoc.xmlPreferences.defaultTableTagName
            //|| TagName == 'topicItem'
            //|| TagName == 'xref'
            )
                added = true;
            
            if (imporCheck) {
                if (!added)
                    xmlMaps.add(TagName, TagName);
                else
                    xmlMaps.add(TagName, None);
            } else {
                if (!added) {
                    for (var j = 2 ; j < allParagraphStyles.length ; j++) {
                        var ps = allParagraphStyles[j].name;
                        if (ps == TagName) {
                            xmlMaps.add(TagName, allParagraphStyles[j]);
                            added = true;
                            break;
                        }
                    }
                }
                
                if (!added) {
                    for (var j = 1 ; j < allCharacterStyles.length ; j++) {
                        var cs = allCharacterStyles[j].name;
                        if (cs == TagName) {
                            xmlMaps.add(TagName, allCharacterStyles[j]);
                            added = true;
                            break;
                        }
                    }
                }
            
                if (!added) {
                    for (var j = 0 ; j < allTableStyles.length ; j++) {
                        var ts = allTableStyles[j].name;
                        if (ts == TagName) {
                            xmlMaps.add(TagName, allTableStyles[j]);
                            added = true;
                            break;
                        }
                    }
                }
            
                if (!added) {
                    for (var j = 0 ; j < allCellStyles.length ; j++) {
                        var cs = allCellStyles[j].name;
                        if (cs == TagName) {
                            xmlMaps.add(TagName, allCellStyles[j]);
                            added = true;
                            break;
                        }
                    }
                }
            
                if (!added) {
                    errorMsgs.push(TagName + ' 태그 매핑 실패');
                    result = false;
                    hasError = true;
                }
            }
            
        }
    
        if (result)
            myDoc.mapXMLTagsToStyles();
        else
            return result;

        var need_id = myDoc.xmlElements[0].evaluateXPathExpression("descendant::*[starts-with(name(), 'Chapter') or starts-with(name(), 'Heading') or name()='Description-NoHTML'][not(@id)]");
        var needs_count = need_id.length;
        for (var i=0 ; i<needs_count ; i++)
            need_id[i].xmlAttributes.add('id', 'd' + need_id[i].id);
            
        var nested = myDoc.xmlElements[0].evaluateXPathExpression("descendant::*[starts-with(name(), 'Empty')][not(@nested)][preceding-sibling::*[not(starts-with(name(), 'Empty'))][1][starts-with(name(), 'OrderList') or starts-with(name(), 'UnorderList')]]");
        var nested_count = nested.length;
        for (var i=0 ; i<nested_count ; i++)
            nested[i].xmlAttributes.add('nested', 'nested');
        
        return result;
    } catch (ex) {
        errorMsgs.push('MappingToStyle 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetMMIattribute(myDoc) {
    try {
        if (mmiStyles.length() > 0) {
            var mmiNames = new Array();
            for (var i = 0 ; i < mmiStyles.length() ; i++) {
                mmiNames.push("name()='" + mmiStyles[i].attribute('name').toString() + "'");
            }
            var myChapterName = myDoc.name.split( '.' )[0];
            var exp = '[' + mmiNames.join(' or ') + ']';
            var mmis = myDoc.xmlElements[0].evaluateXPathExpression('descendant::*' + exp);
            for (var i = 0 ; i < mmis.length ; i++) {
                var mmi = mmis[i];
                var idCount = ++mmiCount;
                try {
                    app.activeDocument;
                    idCount = myChapterName + '_' + idCount.toString();
                } catch (ex) {
                    idCount = idCount.toString();
                }
                mmi.xmlAttributes.add( 'ID', idCount );
                mmi.xmlAttributes.add( 'Chapter', myChapterName );
                /* try {
                    mmi.xmlAttributes.add( 'Page', mmi.paragraphs[0].parentTextFrames[0].parentPage.name );
                } catch (ex) {
                    try {
                        mmi.xmlAttributes.add( 'Page', mmi.paragraphs[0].parentTextFrames[0].parent.name );
                    } catch (ex2) {
                        throw new Error(mmi.contents + ' 의 페이지 정보를 찾을 수 없습니다. 내용이 넘쳐 흐른 텍스트 프레임이 있는지 확인해 주십시오.');
                    }
                } */
            }
        }
    } catch (ex) {
        errorMsgs.push('SetMMIattribute 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}