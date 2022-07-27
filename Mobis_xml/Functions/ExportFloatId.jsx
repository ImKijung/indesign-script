//var root = XML('<root></root>');

function functionExportFloatId() {
    //arguments: handler, mustSinge, needSave, needVisibleOpen
    progressTitle = 'Float Id 추출 중: ';
    var result = function_core(exportFloatId, false, false, false);
    if (hasError || isStopped)
        return false;
    else if (result == null || !result)
         return result;
    else if (typeof(result) == 'xml') {
        //alert($.fileName);
        //alert(result.toXMLString());
        return SaveXMLtoExcel(result, null, null, SetInitial, 2, ['file', 'depth', 'id', 'float', 'string', 'class']);
    } else
        return result;
}

function exportFloatId(myDoc, result) {
    try {
        if (result == null)
            result = XML('<root></root>');
        
        var docName = myDoc.name.replace('.indd', '');
        var story_count = myDoc.stories.length;
        for (var i=0 ; i<story_count ; i++) {
            var paras = myDoc.stories[i].paragraphs;
            var para_count = paras.length;
            for (var j=0 ; j<para_count ; j++) {
                var para = paras[j];
                var con = para.contents;
                var paraStyleName = para.appliedParagraphStyle.name;
                for (var k = 0 ; k < floatIdStyles.length() ; k++) {
                    var floatIdStyleName = floatIdStyles[k].attribute('name').toString();
                    if (paraStyleName == floatIdStyleName) {
                        var depth = para.appliedParagraphStyle.name.replace(/[^0-9]/g,'');
                        if (depth == null || depth == '')
                            depth = '0';
                        var contents = under20(escapeXml(para.contents));
                        result.appendChild(XML('<item file="' + docName + '" '          //A
                                                    + ' depth="' + depth + '" '         //B
                                                    + ' id=""'                          //C
                                                    + ' float=""'	                    //D
                                                    + ' string="' + contents + '"'	    //E
                                                    + ' class="' + paraStyleName + '"'  //F
                                                    + '/>'));
                        break;
                    }
                }
            }
        }
        return result;
    } catch (ex) {
        errorMsgs.push('exportFloatId 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetInitial(vbscript) {
    try {
        vbscript += '\n sheet.Cells(1,1).Value2 = "파일명"';
        vbscript += '\n sheet.Cells(1,2).Value2 = "Depth"';
        vbscript += '\n sheet.Cells(1,3).Value2 = "바로가기 적용 ID"';
        vbscript += '\n sheet.Cells(1,4).Value2 = "플로팅 메뉴 적용"';
        vbscript += '\n sheet.Cells(1,5).Value2 = "contents string"';
        vbscript += '\n sheet.Cells(1,6).Value2 = "Paragraph Style"';
        vbscript += '\n ';
        
        return vbscript;
    } catch (ex) {
        errorMsgs.push('SetInitial 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}
