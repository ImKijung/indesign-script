//function, single document function, need save, open with visible option
function function_core(handler, single, save, visible) {
    try {
        var result = null;
        var documents = select_documents(single);
        if (!documents)
            return documents;
        else if (documents.length > 1) {
            if (single == true) {
                isStopped = true;
                errorMsgs.push('열린 문서가 대상인 기능입니다. 문서 1개가 열린 상태에서 작업해 주십시오.');
                return false;
            } else if (save && check_save() != true) {
                isStopped = true;
                return false;
            }
        }
        
        if (imageSize != null && imageSize != 'UseAttribute') {
            if (parseInt(imageSize) == NaN
                || parseInt(imageSize) == Infinity
                || parseInt(imageSize) == 0) {
                errorMsgs.push('이미지 사이즈가 올바르지 않습니다: ' + imageSize);
                isStopped = true;
                return false;
            }
        }

        for (var i = 0 ; i < documents.length ; i++) {
            if (isStopped) {
                if (documents.length != 1)
                    progresspanel.progressbar.value += 1;
                break;
            }
            progresspanel.text = progressTitle + i + '/' + documents.length;
            var doc = null;
            if (documents.length == 1)
                doc = documents[i];
            else
                doc = app.open(documents[i], visible);

            result = handler(doc, result);
            
            if (documents.length != 1) {
                if (save && result) {
                    var docFile = null;
                    try {
                        docFile = doc.fullName;
                        doc.save();
                    } catch (ex2){
                        try {
                            docFile = File(doc.filePath.fullName + '\\' + doc.name);
                            doc.save(docFile);
                        } catch (ex) {
                            errorMsgs.push(doc.name + ' 문서를 저장할 수 없습니다: Line:' + ex.line + ':: ' + ex);
                            hasError = true;
                        }
                    }
                }
                doc.close();
                progresspanel.progressbar.value += 1;
            }
            progresspanel.text = progressTitle + (i + 1) + '/' + documents.length;
        }
        progresspanel.close();

        return result;
    } catch (ex) {
        errorMsgs.push('function_core 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        CloseAllInvisibleDocuments();
        progresspanel.close();
        return ex;
    }
}

function select_documents(single) {
    try {
        var documents = [];
        if (single != 'multi' && app.documents.length > 0) {
            try {
                documents[0] = app.activeDocument;
            } catch (ex) {
                for (var i = app.documents.length - 1 ; i >= 0 ; i--) {
                    var doc = app.documents[i];
                    if (doc.visible == false)
                        doc.close(SaveOptions.NO);
                }
                return select_documents(single);
            }
        } else if (app.books.length > 0) {
            if (app.books.length > 1) {
                if (askBook) {
                    var sb = null;
                    sb = select_book();
                    if (!sb)
                        return sb;
                    else
                        app.activeBook = sb;
                }
            }
            var book = app.activeBook;
            var contents = book.bookContents;
            //set disable contents codes....
            for (var i = 0 ; i < contents.length ; i++)
                documents[i] = contents[i].fullName;
        } else {
            errorMsgs.push('필요한 문서나 책이 열려있지 않습니다.');
            isStopped = true;
            return false;
        }

        if (documents.length < 1) {
            //errorMsgs.push('select_documents 오류');
            //hasError = true;
            throw new Error('작업대상인 문서나 책이 없습니다.', null, 41);
        } else {
            progresspanel.progressbar.maxvalue = documents.length;
            progresspanel.progressbar.value = 0;
            progresspanel.show();
            progresspanel.active = true;
            return documents;
        }
    } catch (ex) {
        errorMsgs.push('select_documents 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function select_book() {
    try {
        var books = app.books;
        var res =
            "dialog { alignChildren: 'fill', \
                Alert: Panel { orientation: 'column', alignChildren:'left', \
                    text: '다음 BOOK 중에 하나를 선택해주세요'}\
                }, \
                Books: Panel { orientation: 'column', alignChildren:'center', \
                    text: '문서 종류' \
                } \
            }";
        var select = new Window(res, 'BOOK 파일이 여러개 열려있습니다');
        select.hepTip = '목록의 BOOK 중에 하나를 선택합니다';
        var booksCount = books.count();
        var myBook = [];
        for (var i = 0 ; i < booksCount ; i++) {
            myBook[i] = select.children[0].add('button', undefined, GetFileNameWithoutExtention(books[i].name));
            myBook[i].helpTip = i;
            myBook[i].minimumSize = [222, 25]
        };
        select.center();
        var book = null;
        for (var i = 0 ; i < myBook.length ; i++) {
            var n = i;
            myBook[i].onClick = function() {
                try {
                    var index = i;
                    book = books[index];
                    select.close(1);
                } catch (ex) {
                    alert(ex);
                    hasError = true;
                    select.close(0);
                }
            }
        }
        var s = select.show();
        if (s == 1)
            return book;
        else {
            isStopped = true;
            return false;
        }
    } catch (ex) {
        errorMsgs.push('select_book 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function show_text(myDoc, texts) {
    try {
        var story_index = 0;
        var point_index = 0;
        var pre = myDoc.selection[0];
        if (pre != null) {
            story_index = pre.parent.index;
            point_index = pre.insertionPoints[0].index;
        }
        var select = false;
        for (var i = 0 ; i < texts.length ; i++) {
            if (texts[i].parent.index != story_index) {
                texts[i].select();
                select = true;
                break;
            } else if (texts[i].insertionPoints[0].index > point_index) {
                texts[i].select();
                select = true;
                break;
            }
        }
        if (!select)
            texts[0].select();
        myDoc.selection[0].showText();
        return true;
    } catch (ex) {
        errorMsgs.push('show_text 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function check_save() {
    try {
        if (!askSave)
            return true;
        var check = confirm ('문서들이 확인창 없이 자동 저장됩니다.\r\n진행하시겠습니까?', true, '자동저장 확인');
        if (!check)
            isStopped = true;
        return check;
    } catch (ex) {
        errorMsgs.push('check_save 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function find_grep(myDoc, findwhat, appliedpara, appliedchar) {
    try {
        var setResult = grep_setting(findwhat, appliedpara, appliedchar);
        if (setResult == null || !setResult)
            return setResult;
        else
            return myDoc.findGrep();
    } catch (ex) {
        errorMsgs.push('find_grep 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function find_text(myDoc, findwhat, appliedpara, appliedchar) {
    try {
        var setResult = findtext_setting(findwhat, appliedpara, appliedchar);
        if (setResult == null || !setResult)
            return setResult;
        else
            return myDoc.findText();
    } catch (ex) {
        errorMsgs.push('find_text 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function change_grep(myDoc, findwhat, appliedpara, appliedchar, changeto, changepara, changechar) {
    try {
        var setResult = grep_setting(findwhat, appliedpara, appliedchar, changeto, changepara, changechar);
        if (setResult == null || !setResult)
            return setResult;
        else
            return myDoc.changeGrep();
    } catch (ex) {
        errorMsgs.push('change_grep 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function change_text(myDoc, findwhat, appliedpara, appliedchar, changeto, changepara, changechar) {
    try {
        var setResult = findtext_setting(findwhat, appliedpara, appliedchar, changeto, changepara, changechar);
        if (setResult == null || !setResult)
            return setResult;
        else
            return myDoc.changeText();
    } catch (ex) {
        errorMsgs.push('change_text 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function grep_setting(findwhat, appliedpara, appliedchar, changeto, changepara, changechar) {
    try {
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findChangeGrepOptions = null;
        app.findChangeGrepOptions.includeFootnotes = true;
        app.findChangeGrepOptions.kanaSensitive = true;
        app.findChangeGrepOptions.widthSensitive = true;
        if (findwhat != null)
        app.findGrepPreferences.findWhat = findwhat;
        if (appliedpara != null)
            app.findGrepPreferences.appliedParagraphStyle = appliedpara;
        if (appliedchar != null)
            app.findGrepPreferences.appliedCharacterStyle = appliedchar;
        if (changeto != null)
            app.changeGrepPreferences.changeTo = changeto;
        if (changepara != null)
            app.changeGrepPreferences.appliedParagraphStyle = changepara;
        if (changechar != null)
            app.changeGrepPreferences.appliedCharacterStyle = changechar;
        return true;
    } catch (ex) {
        errorMsgs.push('grep_setting 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function findtext_setting(findwhat, appliedpara, appliedchar, changeto, changepara, changechar) {
    try {
        app.findTextPreferences = app.changeTextPreferences = null;
        app.findChangeTextOptions = null;
        app.findChangeTextOptions.includeFootnotes = true;
        app.findChangeTextOptions.kanaSensitive = true;
        app.findChangeTextOptions.widthSensitive = true;
        if (findwhat != null)
        app.findTextPreferences.findWhat = findwhat;
        if (appliedpara != null)
            app.findTextPreferences.appliedParagraphStyle = appliedpara;
        if (appliedchar != null)
            app.findTextPreferences.appliedCharacterStyle = appliedchar;
        if (changeto != null)
            app.changeTextPreferences.changeTo = changeto;
        if (changepara != null)
            app.changeTextPreferences.appliedParagraphStyle = changepara;
        if (changechar != null)
            app.changeTextPreferences.appliedCharacterStyle = changechar;
        return true;
    } catch (ex) {
        errorMsgs.push('findtext_setting 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function InputFolder(title, myDoc, subFolderName) {
    try {
        if (title == null)
            title = '폴더 선택';
        var dirPath = null;
        if (myDoc.saved) {
            dirPath = myDoc.filePath;
            if (subFolderName != null)
                dirPath = dirPath + '/' + subFolderName;
        }
        var folder = new Folder(dirPath).selectDlg(title);
        if (folder == null) {
            isStopped = true;
            return false;
        } else {
            folderPath = folder.fsName;
            return true;
        }
    } catch (ex) {
        errorMsgs.push('InputFolder 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SaveFolder(title, myDoc, subFolderName) {
    try {
        if (title == null)
            title = '저장 폴더 선택';
        var dirPath = null;
        if (myDoc.saved) {
            dirPath = myDoc.filePath;
            if (subFolderName != null)
                dirPath = dirPath + '/' + subFolderName;
        }
        var folder = new Folder(dirPath).selectDlg(title);
        if (folder == null) {
            isStopped = true;
            return false;
        } else {
            folderPath = folder.fsName;
            return true;
        }
    } catch (ex) {
        errorMsgs.push('SaveFolder 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function FolderMatch(myFile) {
    if (myFile.constructor.name == "Folder")
        return true;
    else
        return false;
}

// 사이즈(양의 정수)를 입력받는 함수
function InputSizeData(title, desc) {
    try  {
        //title: image size
        if (title == null)
            title = '% 사이즈 데이터 입력 창'
        //desc: Input Size (170% => 170):
        if (desc == null)
            desc = '사이즈 입력(170% 일 경우 170 입력): '
        var result = prompt(desc, '0', title);
        if (result == null) {
            isStopped = true;
            return false;
        }
        var size = parseInt (result);
        if (isNaN(size)) {
            alert('올바른 숫자가 아닙니다. 양수만 입력해 주십시오.', '입력 오류');
            return InputSizeData (title, desc);
        } else if (size == 0) {
            alert('0이 입력되었습니다. 1 이상의 숫자를 입력해 주십시오.', '입력 오류');
            return InputSizeData (title, desc);
        } else {
            imageSize = size;
            return true;
        }
    } catch (ex) {
        errorMsgs.push('InputSizeData 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function escapeXml(unsafe) {
    return unsafe.replace(/[<>&'"]/g, function (c) {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\'': return '&apos;';
            case '"': return '&quot;';
        }
    });
}

function under20(str) {
    return str.replace(/[\ufeff\r\n]/g, function (c) { return ''; });
}

function SaveXMLtoExcel(xml, file, templateFile, initialHandler, rowNumber, attributes) {
    try {
        if (file == null)
            file = File.saveDialog ('엑셀 파일 저장', 'Excel files:*.xlsx');
        if (file == null) {
            isStopped = true;
            return false;
        }

        var vbscript;
        vbscript = 'returnValue = Main()\n'
                + ' Function Main()\n'
                + '     Dim app, book, sheet\n'
                + '     Set app = CreateObject("Excel.Application")\n'
                + '     On Error Resume Next\n'
                + '     \'take away the \' from the line below if you want to see excel do it\'s stuff\n'
                + '     \'app.visible = true\n'
                + '     app.visible = false\n'
                + '     app.DisplayAlerts = false\n';
        if (templateFile != null && templateFile.exists) {
            vbscript += 'Set book = app.WorkBooks.Open("' + templateFile.fsName + '")\n';
        } else {
            vbscript += 'Set book = app.Workbooks.Add()\n';
        }
        vbscript += '   Set sheet = book.WorkSheets(1)\n';
        vbscript += '   book.SaveAs("' + file.fsName + '")\n';
        
        if (initialHandler != null)
            vbscript = initialHandler(vbscript);
        vbscript += '\n';

        var items = xml.children();
        var numberOfRows = items.length();
        var rowContents = [];
        var setRange = 'sheet.Range("A';
        var openMark = '").Value2 = Array(';
        var closeMark = ')';
        
        var numberOfColumns = attributes.length;
        var toRange = GetLastExcelColumnName (numberOfColumns - 1);
        
        for (var i = 0; i < numberOfRows; i++) {
            var columnContents = [];
            for (var columnNumber = 0; columnNumber < numberOfColumns; columnNumber++) {
                var cellContents = '"' + ReplaceAll(items[i].attribute(attributes[columnNumber]).toString(), '"', '""' ) + '"';
                columnContents .push(cellContents);
            }  
            rowContents[i] = setRange + (i + rowNumber) + ':' + toRange + (i + rowNumber) + openMark + columnContents.join(', ') + closeMark;
        }
        var allContents = rowContents.join('\n') + '\n';
        vbscript += allContents;
        vbscript += '\n sheet.Range("A:' + toRange + '").EntireColumn.Autofit\n'
                + '     book.Save()\n'
                + '     app.ActiveWorkbook.Close\n'
                + '     app.Quit\n'
                + '     Set book = nothing\n'
                + '     Set app = nothing\n'
                + ' If Err.Number <> 0 Then\n'
                + '     Main = Err.Description\n'
                + ' Else\n'
                + '     Main = True\n'
                + ' End If\n'
                + 'End Function\n';
        var value = app.doScript (vbscript, ScriptLanguage.VISUAL_BASIC);
        if (value == true) {
            file.execute();
            return true;
        } else {
            errorMsgs.push(value);
            hasError = true;
            return false;
        }
    } catch (ex) {
        errorMsgs.push('SaveXMLtoExcel 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function GetLastExcelColumnName (columnIndex) {// 0 is A 25 is Z 26 is AA etc.
    // parsed from http://stackoverflow.com/questions/181596/how-to-convert-a-column-number-eg-127-into-an-excel-column-eg-aa
    var dividend = columnIndex + 1;
    var columnName = "";
    var modulo;
  
    while (dividend > 0) {
        modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode (65 + modulo) + columnName;
        dividend = Math.floor((dividend - modulo) / 26);
    }
    return columnName;
} 

function SetPageBreakAttr(myDoc) {
    try {
        grep_setting("~M");
        var pbs = myDoc.findGrep();
        for (var i = 0 ; i < pbs.length ; i++)
            pbs[i].associatedXMLElements[0].xmlAttributes.add('paragraphBreakType', 'NextColumn');
        return true;
    } catch (ex) {
        errorMsgs.push('SetPageBreakAttr 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetPageBreak(myDoc) {
    try {
        var root = myDoc.xmlElements.item(0);
        var pbs = root.evaluateXPathExpression('//*[@paragraphBreakType = \'NextColumn\']');
        
        for (var i = 0 ; i < pbs.length ; i++) {
            pbs[i].insertTextAsContent(SpecialCharacters.COLUMN_BREAK, XMLElementPosition.ELEMENT_END);
            if (pbs[i].characters[pbs[i].characters.length - 2].contents == '\r')
                pbs[i].characters[pbs[i].characters.length - 2].remove();
        }
        return true;
    } catch (ex) {
        errorMsgs.push('SetPageBreak 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function GetFileNameWithoutExtention(fileName) {
    var names = fileName.split('.');
    if (names.length == 1)
        return names[0];
    else {
        names.pop();
        return names.join('.');
    }
}

function ChangeExtention(fileName, newExt) {
    return GetFileNameWithoutExtention(fileName) + '.' + ReplaceAll(newExt, '.', '');
}

function GetParagraphElementFromCharElement(charElement) {
    try {
        if (charElement.markupTag.name == 'Story' || charElement.markupTag.name == 'Root') {
            alert('구조가 잘못되었습니다');
            exit();
        } else if (charElement.parent.markupTag.name == 'Story')
            return charElement;
        else
            return GetParagraphElementFromCharElement(charElement.parent);
    } catch (ex) {
        errorMsgs.push('GetParagraphElementFromCharElement 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function Trim(stringToTrim, token) {
    return TrimEnd(TrimStart(stringToTrim, token), token); //stringToTrim.replace(reg,"");
}

function TrimEnd(stringToTrim, token) {
    if (token == null) {
        token = "\\s+";
    }
    var reg = new RegExp(token + '$', 'g');
    return stringToTrim.replace(reg,"");
}

function TrimStart(stringToTrim, token) {
    if (token == null) {
        token = "\\s+";
    }
    var reg = new RegExp('^' + token, 'g');
    return stringToTrim.replace(reg,"");
}

function GetVisibleDocumentCount() {
    try {
        var visibleDocuments = 0;
        for (var i = 0 ; i < app.documents.length ; i++) {
            var doc = app.documents[i];
            if (doc.visible)
                visibleDocuments += 1;
        }
        return visibleDocuments;
    } catch (ex) {
        errorMsgs.push('GetVisibleDocumentCount 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function CloseAllInvisibleDocuments() {
    try {
        for (var i = 0 ; i < app.documents.count() ; i++) {
            if (!app.documents[i].visible) {
                app.documents[i].close(SaveOptions.NO);
                i = -1;
            }
        }
    } catch (ex) {
        errorMsgs.push('CloseAllInvisibleDocuments 함수 오류: Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function SetLayerVisible(myDoc, layerVisibles, visible) {
    try {
        var layers = myDoc.layers;
        if (visible) {
            for (var i = 0 ; i < layers.length ; i++) {
                layerVisibles[i] = layers[i].visible;
                layers[i].visible = true;
            }
        } else {
            for (var i = 0 ; i < layers.length ; i++) {
                layers[i].visible = layerVisibles[i];
            }
        }
    } catch (ex) {
        errorMsgs.push('SetLayerVisible 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

function ReplaceAll(sValue, param1, param2) {
    return sValue.split(param1).join(param2);
}

function CallDate() {
    var time = new Date();
    var cYear = time.getFullYear();
    var cMonth = PadZero(time.getMonth() + 1, 2);
    var cDate = PadZero(time.getDate(), 2);

    return cYear + '-' + cMonth + '-' + cDate + '_' + ReplaceAll(time.toLocaleTimeString(), ':', '');
}

function MakeZero(timeData) {
    timeData = (timeData >> 1 | timeData >> 2) < 3 ? '0' + timeData.toString() : timeData.toString();
    return timeData;
}

function PadZero(number, length) {
    var str = '' + number;
    while (str.length < length) {
        str = '0' + str;
    }
    return str;
}