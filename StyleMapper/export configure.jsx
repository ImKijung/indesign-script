var xml = XML('<apps></apps>');
var arrayxml = new Array;
if (app.documents.length > 0)
{
    alert('문서를 모두 닫아주세요');
    exit();
}
else if (app.books.length > 0)
{
    try
    {
        var cScript = new File($.fileName);
        var path = cScript.parent.fsName;
        
        var vbscript = '''version()
                            Function version
                                Dim myURL
                                myURL = "http://download.astkorea.net/InDesignScript/Configuration_M3_2015.xlsx"
                                Dim WinHttpReq
                                Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
                                WinHttpReq.setOption 2, 13056
                                WinHttpReq.open "GET", myURL, False
                                WinHttpReq.send ""
                                
                                set objStream = CreateObject("ADODB.Stream")
                                objStream.Type = 1 'adTypeBinary
                                objStream.Open
                                objStream.Write WinHttpReq.responseBody
                                objStream.SaveToFile "''' + path + '''\\Configuration_M3_2015.xlsx", 2
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
                                    Set book = app.WorkBooks.Open("''' + path + '''\\Configuration_M3_2015.xlsx")
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
                else if (name.search('applications') > 1)
                    Applink (mydoc);
                else if (name.search('settings') > 1)
                {
                    Applink (mydoc);
                    vbscript += '\n Set sheet = book.WorkSheets(3)';
                    var array = Category (mydoc);
                    vbscript += '\n sheet.Range("C9:C' + (8+ array.length) + '") = app.Transpose(Array(' + array.join(', ') + '))';
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
            case '\'': return '&apos;';
            case '"': return '&quot;';
        }
    });
}

function under20(str)
{
    return str.replace(/[\ufeff\r\n]/g, function (c) { return ''; } );
}