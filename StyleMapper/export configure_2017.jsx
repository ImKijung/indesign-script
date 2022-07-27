var xml = XML('<apps></apps>');
var xml2 = XML('<apps></apps>');
var arrayxml = new Array;
var arrayxml2 = new Array;
var count = 2;
var chapter = 0;
var hc = 0;

if (app.documents.length > 0)
{
    alert('문서를 모두 닫아주세요');
    exit();
}
else if (app.books.length > 0)
{
    try
    {
        var cScript = app.activeScript;
        //var path = //"C:\\Users\\노건우\\AppData\\Roaming\\Adobe\\InDesign\\Version 12.0-J\\ko_KR\\Scripts\\Scripts Panel";// cScript.parent.fsName;
        var path = cScript.parent.fsName;
        
        var vbscript = ''''version()
                            Function version
                                Dim myURL
                                myURL = "http://download.astkorea.net/InDesignScript/Configuration_M3_2017.xlsx"
                                Dim WinHttpReq
                                Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
                                WinHttpReq.setOption 2, 13056
                                WinHttpReq.open "GET", myURL, False
                                WinHttpReq.send ""
                                
                                set objStream = CreateObject("ADODB.Stream")
                                objStream.Type = 1 'adTypeBinary
                                objStream.Open
                                objStream.Write WinHttpReq.responseBody
                                objStream.SaveToFile "''' + path + '''\\Configuration_M3_2017.xlsx", 2
                                objStream.Close
                                set objStream = Nothing

                            End Function''';
                                
        app.doScript (vbscript, ScriptLanguage.VISUAL_BASIC);  

        var file = File.saveDialog ('저장 위치', 'Excel file:*.xlsx');
        if (file == null)
        {
            alert('취소하였습니다');
            exit();
        }
        else if (file.exists)
        {
            if (!confirm('파일이 이미 있습니다.\r\n기존 파일을 바꾸시겠습니까?', true, '경고'))
                exit();
        }

        var xml = XML('<apps></apps>');
        var book = app.activeBook;
        var bookcontents = book.bookContents;
        var contents_count = bookcontents.length;

        var vbscript = '''Main()
                                Sub Main()
                                    Dim app, book, sheet
                                    Set app = CreateObject("Excel.Application")
                                    app.visible = false
                                    app.DisplayAlerts = false
                                    Set book = app.WorkBooks.Open("''' + path + '''\\Configuration_M3_2017.xlsx")
                                    Set sheet = book.WorkSheets(1)''';
        vbscript +='\n book.SaveAs("' + file.fsName + '")';
        var script = null;
        
        var rownumber = 13;
        for (var i=0 ; i<contents_count ; i++)
        {
            var name = bookcontents[i].name.toLowerCase();
            var realname = bookcontents[i].name;
            if (name.search('cover') < 0 && name.search('_toc') < 0)
            {
                vbscript += '\n Set sheet = book.WorkSheets(1)';
                vbscript += '\n sheet.Cells(' + rownumber++ + ', 3).Value2 = "' + realname.replace('.indd', '.idml') + '"';
            }
            try
            {
                var mydoc = app.open(bookcontents[i].fullName, false);
                if (name.search('cover') > 1)
                {
                    vbscript += '\n Set sheet = book.WorkSheets(1)';
                    var model = Config(mydoc);
                    if (model != null)
                    {
                        vbscript += '\n sheet.Cells(9,3).Value2 = "' + model.join('<br/>') + '"';
                        vbscript += '\n sheet.Cells(9,3).WrapText = true';
                    }
                }
                //else if (name.search('applications') > 1)
                    //Applink (mydoc);
                else if (name.search('settings') > 1)
                {
                    //Applink (mydoc);
                    vbscript += '\n Set sheet = book.WorkSheets(3)';
                    var array = Category (mydoc);
                    vbscript += '\n sheet.Range("C9:C' + (8+ array.length) + '") = app.Transpose(Array(' + array.join(', ') + '))';
                }
                if ( i > 1 && i < 7) {
                    Applink (mydoc);
                    Headings(mydoc);
                }
            }
            catch(ex)
            {
                alert(ex);
            }
            finally
            {
                mydoc.close();
            }
        }
        
        if (arrayxml.length > 0)
        {
            vbscript += '\n Set sheet = book.WorkSheets(2)';
            for (var i=0 ; i<arrayxml.length ; i++)
                vbscript += '\n sheet.Range("B' + (i+9) + ':F' + (i + 9) + '").Value2 = Array(' + arrayxml[i].join(', ') + ')';
        }
        if (arrayxml2.length > 0)
        {
            vbscript += '\n Set sheet = book.WorkSheets(6)';
            for (var i=0 ; i<arrayxml2.length ; i++)
                vbscript += '\n sheet.Range("B' + (i+9) + ':J' + (i + 9) + '").Value2 = Array(' + arrayxml2[i].join(', ') + ')';
        }
        try
        {
            vbscript +=  '''\n book.Save()
                app.ActiveWorkbook.Close
                app.Quit 
                Set book = nothing
                Set app = nothing
            End Sub''';
            
            if (file.exists)
            {
                try
                {
                    var b = file.remove();
                    if (!b)
                    {
                        alert('기존 파일을 바꿀 수 없습니다', '오류');
                        exit();
                    }
                }
                catch(ex)
                {
                    alert(ex);
                    exit();
                }
            }
            vbscript = vbscript.replace(/&apos;/gi, "'");
            app.doScript (vbscript, ScriptLanguage.VISUAL_BASIC);  
            alert('완료되었습니다', '완료');
            var path = new Folder(file.path);
            path.execute();
        }
        catch(ex)
        {
        }
    }
    catch (ex)
    {
        alert(ex);
    }
}
else
{
    alert('북 파일이 없습니다');
    exit();
}

function Config(mydoc)
{
    var rName = new Array;
    var lName = new Array;
    var story_count = mydoc.stories.length;
    for (var i=0 ; i<story_count ; i++)
    {
        var paras = mydoc.stories[i].paragraphs;
        var para_count = paras.length;
        for (var j=0 ; j<para_count ; j++)
        {
            if (paras[j].appliedParagraphStyle.name == 'ModelName-Cover')
                rName.push(under20(paras[j].contents.replace ('\r', '').replace ('\n', '')));
            else if (paras[j].appliedParagraphStyle.name == 'ModelName-Cover-Left')
                lName.push(under20(paras[j].contents.replace ('\r', '').replace ('\n', '')));
        }
    }
    if (lName.length > 0)
        return lName;
    else if (rName.length > 0)
        return rName;
    else
        return null;
}

function Applink(mydoc)
{
    var story_count = mydoc.stories.length;
    for (var i=0 ; i<story_count ; i++)
    {
        var paras = mydoc.stories[i].paragraphs;
        var para_count = paras.length;
        for (var j=0 ; j<para_count ; j++)
        {
            if (paras[j].appliedParagraphStyle.name.search('-APPLINK') > 1)
            {
                xml.appendChild(XML('<app file="' + mydoc.name.replace('.indd', '.idml') + '" depth="' + paras[j].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '" id="" image_name=""  string="' + under20(escapeXml(paras[j].contents)) + '"/>'));
                arrayxml.push(Array('"' + mydoc.name.replace('.indd', '.idml') + '"',  '"' + paras[j].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '"', '""', '""',  '"' + under20(escapeXml(paras[j].contents)) + '"'));
            }
        }
    }
}
function Headings(mydoc)
{
    var textDestinations = mydoc.hyperlinkTextDestinations;
    var story_count = mydoc.stories.length;
    var h1 = 0;
    var h2 = 0;
    var h3 = 0;
    var h4 = 0;
    for (var i=0 ; i<story_count ; i++)
    {
        var paras = mydoc.stories[i].paragraphs;
        var para_count = paras.length;
        var no = '';
        var link = '';
        for (var j=0 ; j<para_count ; j++)
        {
            if (paras[j].appliedParagraphStyle.name.search('Heading') > -1) {
                if (paras[j].appliedParagraphStyle.name.search('Heading1') > -1) {
                    h1++;
                    h2 = 0;
                    h3 = 0;
                    h4 = 0;
                    no = 'h1_' + hc++;
                    link = mydoc.name.replace('.indd', '').toLowerCase() + '_' + h1 + '.html';
                } else if (paras[j].appliedParagraphStyle.name.search('Heading2') > -1) {
                    h2++;
                    h3 = 0;
                    h4 = 0;
                    no = count;
                    link = mydoc.name.replace('.indd', '').toLowerCase() + '_' + h1 + '.html#' + (h2-1);
                } else if (paras[j].appliedParagraphStyle.name.search('Heading3') > -1) {
                    var dummy2 = h2 == 0 ? 0 : h2 - 1;
                    h3++;
                    h4 = 0;
                    no = count;
                    link = mydoc.name.replace('.indd', '').toLowerCase() + '_' + h1 + '.html#' + (dummy2) + '#' + (h3-1);
                } else if (paras[j].appliedParagraphStyle.name.search('Heading4') > -1) {
                    var dummy2 = h2 == 0 ? 0 : h2 - 1;
                    var dummy3 = h3 == 0 ? 0 : h3 - 1;
                    h4++;
                    no = count;
                    link = mydoc.name.replace('.indd', '').toLowerCase() + '_' + h1 + '.html#' + (dummy2) + '#' + (dummy3) + '#' + (h4-1);
                }
                count++;
                xml2.appendChild(XML('<app file="' + link + '" '	//B
                                            + ' depth="' + paras[j].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '" ' 	//C
                                            + ' id="" icon_name=""'	//D, E
                                            + ' topic_icon_name=""'	//F
                                            + ' string="' + under20(escapeXml(paras[j].contents)) + '"'	//G
                                            + ' class="' + paras[j].appliedParagraphStyle.name + '"'		//H
                                            + '/>'));
                arrayxml2.push(Array('"' + link + '"',
                                            '"' + paras[j].appliedParagraphStyle.name + '"', '"' + paras[j].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '"',
                                            '"' + no + '"',  '"' + chapter + '"', '""', '""', '""',  '"' + under20(escapeXml(paras[j].contents)) + '"'));
            } else if (paras[j].appliedParagraphStyle.name.search('Chapter') > -1) {
                h1 = 0;
                h2 = 0;
                h3 = 0;
                h4 = 0;
                no = count;
                link = mydoc.name.replace('.indd', '').toLowerCase();
                chapter++;
                count++;
                xml2.appendChild(XML('<app file="' + link + '" '	//B
                                            + ' depth="' + paras[j].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '" ' 	//C
                                            + ' id="" icon_name=""'	//D, E
                                            + ' topic_icon_name=""'	//F
                                            + ' string="' + under20(escapeXml(paras[j].contents)) + '"'	//G
                                            + ' class="' + paras[j].appliedParagraphStyle.name + '"'		//H
                                            + '/>'));
                arrayxml2.push(Array('"' + link + '"',
                                            '"' + paras[j].appliedParagraphStyle.name + '"', '"' + paras[j].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '"',
                                            '"' + no + '"',  '"' + chapter + '"', '""', '""', '""',  '"' + under20(escapeXml(paras[j].contents)) + '"'));
            } else if (paras[j].appliedParagraphStyle.name.search('Empty') > -1) {
                var tables = paras[j].tables;
                for (var k=0 ; k<tables.length ; k++) {
                    var table = tables[k];
                    var cells = table.cells;
                    for (var l=0 ; l<cells.length ; l++) {
                        var cell = cells[l];
                        var pic = cell.paragraphs;
                        for (var m=0 ; m<pic.length ; m++) {
                            if (pic[m].appliedParagraphStyle.name.search('Heading3') > -1) {
                                var dummy2 = h2 == 0 ? 0 : h2 - 1;
                                h3++;
                                h4 = 0;
                                no = count++;
                                link = mydoc.name.replace('.indd', '').toLowerCase() + '_' + h1 + '.html#' + (dummy2) + '#' + (h3-1);
                                arrayxml2.push(Array('"' + link + '"',
                                                            '"' + pic[m].appliedParagraphStyle.name + '"', '"' + pic[m].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '"',
                                                            '"' + no + '"',  '"' + chapter + '"', '""', '""', '""',  '"' + under20(escapeXml(pic[m].contents)) + '"'));
                            } else if (pic[m].appliedParagraphStyle.name.search('Heading4') > -1) {
                                var dummy2 = h2 == 0 ? 0 : h2 - 1;
                                var dummy3 = h3 == 0 ? 0 : h3 - 1;
                                h4++;
                                no = count++;
                                link = mydoc.name.replace('.indd', '').toLowerCase() + '_' + h1 + '.html#' + (dummy2) + '#' + (dummy3) + '#' + (h4-1);
                                arrayxml2.push(Array('"' + link + '"',
                                                            '"' + pic[m].appliedParagraphStyle.name + '"', '"' + pic[m].appliedParagraphStyle.name.replace(/[^0-9]/g,'') + '"',
                                                            '"' + no + '"',  '"' + chapter + '"', '""', '""', '""',  '"' + under20(escapeXml(pic[m].contents)) + '"'));
                            }
                        }
                    }
                }
            }
        }
    }
}

function Category(mydoc)
{
    var array = [];
    var story_count = mydoc.stories.length;
    for (var i=0 ; i<story_count ; i++)
    {
        var paras = mydoc.stories[i].paragraphs;
        var para_count = paras.length;
        for (var j=0 ; j<para_count ; j++)
        {
            if (paras[j].appliedParagraphStyle.name.search('Heading1') > -1 ) // && paras[j].appliedParagraphStyle.name.search('-Intro') < 0)
                array.push ('"' + under20(paras[j].contents) + '"');
        }
    }
    return array;
}

function escapeXml(unsafe) {
    return unsafe.replace(/[<>&'"]/g, function (c) {
        switch (c) {
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '&': return '&amp;';
            case '\\': return '&bsol;';
            case '"': return '&quot;';
            case "'": return '&apos;';
        }
    });
}

function under20(str)
{
    return str.replace(/[\ufeff\r\n]/g, function (c) { return ''; } );
}