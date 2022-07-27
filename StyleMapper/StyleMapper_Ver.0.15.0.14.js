// by 건우
// 로그는 아래에

var check_CharacterError = 'CharacterStyle Error Log' + '\r\n\r\n';
var write_ErrorAddress = '';
var errorCount = 0;
var indexError = 0;
var totalError = 0;
var condition = 0;
var resultLog = '에 실패한 Indd 파일\r\n';
var indexingLog = '색인 작업이 실패한 Indd 파일';
var styleErrorCount = 0;
var totalXMLCount = 0;
var IndexCondition = 0;
var myDocNumber = 1;
var MMIcounter = 0;
var no_intergrate = false;
var integrate = false;
var myNone = app.characterStyles[0].name;
var Cell_Table_H_Tint;
var Cell_Table_H_K_Tint;
var Cell_Number_Tint;
var Temp_XML = new Array();
var qsg_type = false;
var export_all = false;
var global_hyperlink = new Array();
var global_destination = new Array();
var check_stop = false;
var myXSLT = null;
var linked = [];

//Define Colors
var colorC_Image = [0, 54, 30, 0];  //홍시
var colorChangedMMI = [55, 0, 75, 0];   //연두
var colorChangedMMI2 = [0, 44, 91, 0];  //주황
var colorMMI = [0, 38, 76, 0];   //연두
var colorMMI_NoBold = [0, 38, 76, 0];   //연두CHk
var colorIword = [16, 0, 76, 0];    //test
var colorIword_NoBold = [26, 0, 76, 0]; //연두
var colorC_For_English = [0, 44, 91, 0];    //주황
var colorC_FontChange = colorC_For_English;
var colorC_Singlestep = [16, 85, 0, 0]; //자주
var colorC_SingleStep = [16, 85, 0, 0]; //자주
var colorC_NoBreak = [16, 85, 0, 0];
var Wrong_Color = [16, 85, 0, 0];   //자주

var colorCell = [0, 100, 100, 0];
var colorCell_Table_H = [55, 15, 0, 0];
var colorCell_TableHead_Fill = [55, 15, 0, 0];
var colorCell_TableHead_Icon_Fill = [55, 15, 0, 0];
var colorCell_Table_H_K = [60, 0, 9, 0];//바다
var colorCell_TableHead_Empty = [60, 0, 9, 0];
var colorCell_Number = [15, 19, 0, 0];//연보라
var color_app = [0, 40, 20, 0];
var color_recommend = [35, 40, 0, 0];
var color_mmi = [100, 0, 100, 40];

var THICK = '7mm';//밑줄 두께
var ChangedTHICK = '5mm';//밑줄 두께
var MMI_Changed;
var SHIFT = '3pt';//베이스라인 높이

var doNotCheckErrors = false;
var trimCharacterStyleContents = false;

var onlyDocument = false;
var my_charstyle = "C_Image";
var progresspanel = new Window( "window", "실행중" );
with ( progresspanel ) { progresspanel.progressbar = add( "progressbar", [12, 12, 400, 24], 0, 100 ); }
var sFolder = null;
var sFolder_fs = null;
try
{
    var cScript = new File($.fileName);
    var path = cScript.parent.fsName + '\\';
    sFolder = new Folder( cScript.path );
    sFolder_fs = path;

    var vbscript = '''version()
                            Function version
                                Dim myURL
                                myURL = "http://download.astkorea.net/InDesignScript/Script_Version.txt"
                                Dim WinHttpReq
                                Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
                                WinHttpReq.open "GET", myURL, False
                                WinHttpReq.send ""
                                version = WinHttpReq.ResponseText
                                
                                Set objFS = CreateObject("Scripting.FileSystemObject")
                                strTemp = "''' + path + '''Script_Version.txt"
                                Set objOutFile = objFS.CreateTextFile(strTemp,True) 
                                objOutFile.Write(version)
                                objOutFile.Close
                            End Function''';
                                
    app.doScript (vbscript, ScriptLanguage.VISUAL_BASIC);  

    var sVersionFile = new File( cScript.parent.fullName + '/Script_Version.txt')
    if ( !sVersionFile.exists )
    {
        alert( '정상적인 접근이 아닙니다' );
        exit();
    }
    sVersionFile.open( 'r' );
    var sVersion = sVersionFile.read();
    sVersionFile.close();

    if ( cScript.name != sVersion )
    {
        if ( confirm( '구 버전의 스크립트입니다\r\n새 버전의 스크립트를 복사하시겠습니까?\r\n현재버전 : ' + cScript.name + '\r\n신규버전 : '  + sVersion) )
        {
            //var nVersionFile = new File( '//10.10.10.4/strategywork/Mobile/Program/Script/' + sVersion )
            //if ( !nVersionFile.exists )
            //    throw err;
            //nVersionFile.copy( cScript.path + '/' + sVersion );
            vbscript = '''version()
                            Function version
                                Dim myURL
                                myURL = "http://download.astkorea.net/InDesignScript/''' + sVersion + '''"
                                Dim WinHttpReq
                                Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
                                WinHttpReq.setOption 2, 13056
                                WinHttpReq.open "GET", myURL, False
                                WinHttpReq.send ""
                                
                                set objStream = CreateObject("ADODB.Stream")
                                objStream.Type = 1 'adTypeBinary
                                objStream.Open
                                objStream.Write WinHttpReq.responseBody
                                objStream.SaveToFile "''' + path + '''\\''' + sVersion +'''", 2
                                objStream.Close
                                set objStream = Nothing

                            End Function''';
            app.doScript (vbscript, ScriptLanguage.VISUAL_BASIC); 

            vbscript = '''version()
                        Function version
                            Dim myURL
                            myURL = "http://download.astkorea.net/InDesignScript/indent_XML.xsl"
                            Dim WinHttpReq
                            Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
                            WinHttpReq.setOption 2, 13056
                            WinHttpReq.open "GET", myURL, False
                            WinHttpReq.send ""
                            
                            set objStream = CreateObject("ADODB.Stream")
                            objStream.Type = 1 'adTypeBinary
                            objStream.Open
                            objStream.Write WinHttpReq.responseBody
                            objStream.SaveToFile "''' + path + '''\\indent_XML.xsl", 2
                            objStream.Close
                            set objStream = Nothing

                        End Function''';
            app.doScript (vbscript, ScriptLanguage.VISUAL_BASIC); 

            sFolder.execute();
            exit();
        }
    }
}
catch ( err )
{ alert( '서버의 파일버전 정보가 손상되었습니다\r\n현 버전의 스크립트를 실행합니다' + err ); }

do
{
    var res =

        "dialog { alignChildren: 'fill', \
            pn1: Panel { alignChildren:'', \
                text: '메뉴얼 선택', \
                group: Group { orientation: 'row', alignment: 'center', \
                    btnXML: Button { text: '일반 메뉴얼', minimumSize: [95, 50] }, \
                    btnIDML: Button { text: '통합 메뉴얼', minimumSize: [95, 50] },\
                } \
            }, \
            pn2: Panel {  alignChildren:'center', \
                text: '서브 메뉴',\
                group: Group { orientation: 'row', alignment: 'center', \
                    btnSide: Button { text:'부가기능', minimumSize: [95, 50]}, \
                    btnShow: Button { text:' 오류보기', minimumSize: [95, 50]}, \
                }\
            } \
        }";
    selectTool = new Window( res, 'Select Function' );
    selectTool.helpTip = 'XML에 관련된 기능과 이미지, 에러에 관련된 기능 들이 있습니다';
    selectTool.pn1.group.btnXML.helpTip = '일반 메뉴얼을 작업할 때 쓰이는 메뉴입니다';
    selectTool.pn1.group.btnIDML.helpTip = '통합 메뉴얼을 작업할 때 쓰이는 메뉴입니다';
    selectTool.pn2.group.btnSide.helpTip = '오류를 찾거나 인덱스, 이미지, 하이라이트를\r\n 작업할 때 쓰이는 메뉴입니다';
    selectTool.pn2.group.btnShow.helpTip = '에러가 발생한 곳의 구조 요소를 보여줍니다\r\n반드시 로그 파일이 존재해야합니다';
    selectTool.pn1.group.btnXML.onClick = function () { selectTool.close( 3 ); }
    selectTool.pn1.group.btnIDML.onClick = function () { selectTool.close( 4 ); }
    selectTool.pn2.group.btnSide.onClick = function () { selectTool.close( 5 ); }
    selectTool.pn2.group.btnShow.onClick = function () { selectTool.close( 6 ); }
    intResult = selectTool.show() - 2;
    condition = 1;

    if ( intResult )
    {
        switch ( intResult )
        {
            case 1:
                condition = EditXML( 0 );
                break;
            case 2:
                condition = EditIntegrate();
                break;
            case 3:
                condition = mainSide();
                break;
            default:
                condition = ShowErrors();
        }
    }
} while ( !condition )
// 본문 끝-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

//XML작업을 수행하는 기능의 메인부분
function EditXML( XMLToolStatus )
{
    var myImageSize = null;
    var docNames = [];
    var sourceFile = 0;
    var indexResult = null;
    var select_xmlstructure = 0;
    var returnValue = null;
    no_intergrate = true;
    myXSLT = false;

    do
    {
        if ( XMLToolStatus == null || XMLToolStatus == 0 )
            XMLToolStatus = SelectXMLTools();

        if ( XMLToolStatus < 0 )  // 오류
            return -1;
        else if ( XMLToolStatus == null || XMLToolStatus == 0 )    // 취소
            return 0;

        returnValue = Select_Contents( false );
        if ( returnValue < 0 )
            return -1;
        else if ( returnValue == 0 )
            return 0;
        else if ( returnValue == 'continue' )
        {
            XMLToolStatus = 0;
            continue;
        }
        docNames = returnValue;

        switch ( XMLToolStatus )
        {
            case 1:
                myImageSize = SelectManualType();
                if ( myImageSize < 0 )
                    return -1;
                else if ( !myImageSize )
                {    // 취소
                    if ( integrate )
                        return 0;
                    else
                    {
                        XMLToolStatus = 0;
                        continue;
                    }
                }
                resultLog = 'XML 호출' + resultLog;
                if ( myImageSize == 'UM' )
                {
                    indexResult = IndexFunctionSelect();
                    if ( indexResult < 0 )
                        return -1;
                    else if ( indexResult )
                    {
                        sourceFile = SelectSourceINDD( docNames[0] );
                        if ( sourceFile < 0 )
                            return -1;
                    }
                }
                break;
            case 2:
                resultLog = 'XML 추출' + resultLog;
                returnValue = Select_XmlStructure();
                if ( returnValue < 0 )
                    return -1;
                else if ( returnValue == 0 )
                    return 0;
                else if ( returnValue == 'continue' )
                {
                    XMLToolStatus = 0;
                    continue;
                }
                select_xmlstructure = returnValue;
                break;
            case 3:
                resultLog = '오류가 발생한 INDD 파일\r\n';
                returnValue = Select_XmlStructure();
                if ( returnValue < 0 )
                    return -1;
                else if ( returnValue == 0 )
                    return 0;
                else if ( returnValue == 'continue' )
                {
                    XMLToolStatus = 0;
                    continue;
                }
                select_xmlstructure = returnValue;
                break;
            case 5:
                integrate = true;
                alert( Make_Image(), 'C_Image 생성 결과' );
                integrate = false;
                return 1;
                break;
            case 6:
                alert( Delete_Image(), 'C_Image 삭제 결과' );
                return 1;
                break;
            case 7:
                var my_Queries = [
                    ["grep", "NoteToCellBody", myNone],
                    ["grep", "Refer_CharStyle_del", "MMI_NoBold"],
                    ["grep", "Refer_CharStyle_del2", "MMI"],
                    ["grep", "Refer_CharStyle_del8", "C_NoBreak"],
                    ["grep", "del_mmi_nextcolumn", "MMI"],
                    ["grep", "del_mmi_nobold_nextcolumn", "MMI_NoBold"]
                ];
                alert( Redefine_Style( my_Queries ), '참조 스타일 삭제 결과' );
                return 1;
                break;
            case 8:
                ExportIDML();
                return 1;
                break;
            default:
                break;
        }
    } while ( XMLToolStatus == 0 )

    var my_Queries = null;
    if ( qsg_type )
    {
        my_Queries = [
            ["grep", "Refer_CharStyle_del", "MMI_NoBold"],
            ["grep", "Refer_CharStyle_del2", "MMI"],
            ["grep", "Refer_CharStyle_del3", "C_Section_Black"],
            ["grep", "Refer_CharStyle_del4", "C_Section_Blue"],
            ["grep", "Refer_CharStyle_del5", "C_Section_Green"],
            ["grep", "Refer_CharStyle_del6", "C_Section_Red"],
            ["grep", "Refer_CharStyle_del7", "MMI_NoBold"],
            ["grep", "Refer_CharStyle_del8", "C_NoBreak"],
            ["grep", "del_mmi_nextcolumn", "MMI"],
            ["grep", "del_mmi_nobold_nextcolumn", "MMI_NoBold"],
            ["grep", "del_blank_cellbody_bulleted_common", "CellBody_Bullet"],
            ["grep", "del_blank_cellbody_common", "CellBody"],
            ["grep", "del_blank_cellbody_num_common", "CellBody_Numbered"],
            ["grep", "del_blank_cellbody_num1_common", "CellBody_Numbered1"],
            ["grep", "del_blank_note_bulleted_common", "Note_Bulleted"],
            ["grep", "del_blank_note_common", "Note"]
        ];
    }
    else
    {
        my_Queries = [
            ["grep", "Refer_CharStyle_del", "MMI_NoBold"],
            ["grep", "Refer_CharStyle_del2", "MMI"],
            ["grep", "Refer_CharStyle_del3", "C_Section_Black"],
            ["grep", "Refer_CharStyle_del4", "C_Section_Blue"],
            ["grep", "Refer_CharStyle_del5", "C_Section_Green"],
            ["grep", "Refer_CharStyle_del6", "C_Section_Red"],
            ["grep", "Refer_CharStyle_del7", "MMI_NoBold"],
            ["grep", "Refer_CharStyle_del8", "C_NoBreak"],
            ["grep", "del_mmi_nextcolumn", "MMI"],
            ["grep", "del_mmi_nobold_nextcolumn", "MMI_NoBold"],
            ["grep", "del_blank_after_sentence_Common", myNone],
            ["grep", "del_blank_after_table_Common", myNone],
            ["grep", "del_blank_cellbody_bulleted_common", "CellBody_Bullet"],
            ["grep", "del_blank_cellbody_common", "CellBody"],
            ["grep", "del_blank_cellbody_num_common", "CellBody_Numbered"],
            ["grep", "del_blank_cellbody_num1_common", "CellBody_Numbered1"],
            ["grep", "del_blank_note_bulleted_common", "Note_Bulleted"],
            ["grep", "del_blank_note_common", "Note"]
        ];
    }

    returnValue = Working_Step( docNames, my_Queries, XMLToolStatus, select_xmlstructure, myImageSize, sourceFile );
    //if ( returnValue < 0 )
        //return returnValue;

    if ( totalError )
    {
        if ( check_stop )
            alert( '중지 되었습니다', '중지' )
        else if ( XMLToolStatus == 3 )
            alert( resultLog + '\r\n\r\n로그를 확인해 주세요', 'Error Check 결과' );
        else if ( XMLToolStatus == 2 )
            alert( resultLog + '\r\n\r\n로그를 확인해 주세요', 'ExportXML 결과' );
        else if ( IndexCondition )
            alert( resultLog + '\r\n\r\n' + indexingLog + '\r\n\r\n로그를 확인해 주세요', 'ImportXML 결과' );
        else if ( totalError - indexError )
            alert( resultLog + '\r\n\r\n로그를 확인해 주세요', 'ImportXML 결과' );
        else
            alert( '모든 과정이 성공적으로 끝났습니다', 'XML 결과' );
    }
    else
    {
        if ( check_stop )
            alert( '중지 되었습니다', '중지' )
        else if ( XMLToolStatus < 4 )
            alert( '모든 과정이 성공적으로 끝났습니다', 'XML 결과' );
        else
            alert( '태그삭제 완료', 'XML 결과' );
    }
    return 1;
}

function Select_Contents( allDocument )
{
    var docNames = [];
    if ( app.documents.count() > 0 )
    {
        // 문서
        onlyDocument = true;
        var myDoc = app.activeDocument;
        try {
            docNames[0] = myDoc.fullName;
            myDoc.filePath;
        }
        catch ( err )
        {
            alert( myDoc.name + '가 한 번도 저장되지 않았습니다.\r\n저장하신 후에 진행해주세요', '오류' );
            return -1;
        }
    }
    else
    {
        // Book
        onlyDocument = false;
        docNames = SelectTarget( allDocument );
        if ( !docNames )
            return 0;
        else if ( docNames < 0 )
            return -1;

        var res =
        "dialog { alignChildren: 'fill', \
            Success: Panel { orientation: 'row', alignChildren:'center', \
                text: '문서들이 확인창 없이 자동저장 됩니다', \
                s: StaticText { text: '진행하시겠습니까?' },\
            }, \
            buttons: Group { orientation: 'row', alignment: 'right', \
                buttonsC: Group { orientation: 'row', alignment: 'left', \
                    cancelBtn: Button { text:'뒤로', properties:{name:'cancel'} } \
                }\
                buttonsE: Group { orientation: 'row', alignment: 'left', \
                    EmptyBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                }\
                buttonsN: Group { orientation: 'row', alignment: 'right', \
                    okBtn: Button { text:'다음', properties:{name:'ok'} }, \
                } \
            } \
        }";
        var select = new Window( res, '문서 자동저장 경고' );
        select.hepTip = '지정한 문서들이 작업이 끝난 후에 자동으로 저장이 됩니다';
        select.center();
        select.buttons.buttonsE.EmptyBtn.active = false;
        select.buttons.buttonsE.EmptyBtn.visible = false;
        var result = select.show() - 1;

        if ( result > 0 )
        {
            if ( integrate )
                return 0;
            XMLToolStatus = 0;
            return 'continue';
        }
    }
    return docNames;
}

function Select_XmlStructure()
{
    var select_xmlstructure = SelectXMLStructure();
    if ( select_xmlstructure < 0 )
        return -1;
    else if ( select_xmlstructure == 0 )
    {
        if ( integrate )
            return 0;
        XMLToolStatus = 0;
        return 'continue';
    }

    if ( select_xmlstructure == 2 || select_xmlstructure == 4 )
    {
        // MMI Counter
        myXSLT = new File( 'C:/Users/Administrator/AppData/Roaming/Adobe/InDesign/Version 5.0-KR/Scripts/Scripts Panel/MMIFilter.xsl' );
        if ( !myXSLT.exists )
            myXSLT = File.openDialog( 'Select MMIFilter File', 'MMIFilter Files:' + 'MMIFilter.xsl' );
    }
    else
        myXSLT = false;
    if ( select_xmlstructure > 2 )
    {
        select.hepTip = '지정한 문서들이 작업 전, 후에 자동으로 저장이 됩니다.\r\n오류가 없는지 확인 후에 진행해주세요';
        var result = select.show() - 1;
        if ( result > 0 )
            return -1;
    }
    return select_xmlstructure;
}

function Check_XMLfile(bookDocuments)
{
    try
    {
        var error = '';
        var bookDocumentsLength = bookDocuments.length;
        for (var i = 0 ; i < bookDocumentsLength ; i++)
        {
            var xmlfile = bookDocuments[i].name.substring(0, bookDocuments[i].name.lastIndexOf('.')) + '.xml';
             if (!File(bookDocuments[i].filePath + '/XML/' + xmlfile).exists)
             {
                 error += xmlfile + ' 파일이 없습니다.\r\n';
                 totalError++;
             }
         }
        if (error !='')
            return error;
        else
            return true;
    }
    catch (err)
    {
        totalError++;
        return err.toString();
    }
}

// 3 단계
function Working_Step( docNames, my_Queries, XMLToolStatus, XMLStructureStatus, myImageSize, sourceFile )
{
    var bookDocuments = [];
    var bookDocumentsLength = docNames.length;
    var log = null;
    var images = null;
    var imageCount = 0;
    var returnValue = null;
    var docPath= null;
    progresspanel.progressbar.maxvalue = bookDocumentsLength;
    progresspanel.show();
    
    if (!onlyDocument)
    {
        try
        {
            for (var i = 0 ; i < bookDocumentsLength ; i++)
                bookDocuments[i] = app.open(docNames[i], false);
        }
        catch (err)
        {
            var docs = app.documents.count();
            for (var i = 0 ; i < docs ; i++)
                app.documents[0].close(SaveOptions.NO);
            return -1;
        }
    }
    else
    {
        bookDocuments[0] = app.open(docNames[0], false);
        docPath = XML('<docList></docList>');
        for ( var i = 0 ; i < app.documents.count() ; i++)
            docPath.appendChild(XML('<doc path="' + app.documents[i].fullName + '"/>'));
    }    

    switch ( XMLToolStatus )
    {
        case 1: // Import XML
            if (onlyDocument)
                Save_Destination(bookDocuments[0]);
            returnValue = Check_XMLfile(bookDocuments);
            if (returnValue != true)
            {
                resultLog = returnValue;
                return returnValue;
            }
            returnValue = BookContents( bookDocuments, sourceFile, XMLToolStatus, XMLStructureStatus, myImageSize );
            if ( returnValue < 0 )
                return returnValue;
            if (onlyDocument)
                Reconnect_Destination(bookDocuments[0]);
            for ( var i = 0 ; i < bookDocumentsLength ; i++ )
                Set_HyperlinkInform( bookDocuments[i] );
            for ( var i = 0 ; i < bookDocumentsLength ; i++ )
                DeleteTag( bookDocuments[i] );
            break;
        case 2: // Export XML
        case 3: // Check XML
            var log = null;
            var images = null;
            var imageCount = 0;
            for ( var i = 0 ; i < bookDocumentsLength ; i++ )
            {
                images = bookDocuments[i].allGraphics;
                imageCount = images.length;
                for ( var j = 0; j < imageCount ; j++ )
                {
                    try
                    {
                        if ( images[j].itemLink.status == 1819109747 )	// 링크 오류 상태
                        {
                            progresspanel.hide();
                            alert( bookDocuments[i].name.split( '.' )[0] + '문서에 이미지링크 오류가 있습니다.\r\n 링크를 재설정 해주세요', '링크 오류' );
                            totalError++;
                            if (!onlyDocument)
                                bookDocuments[i].close();
                            return -1;
                        }
                    } catch ( err ) { }
                }
                bookDocuments[i] = app.open( docNames[i], true );
                log = RedefineStyles( my_Queries, bookDocuments[i], false );
                if ( errorCount )
                {
                    resultLog += '\r\n' + log;
                    errorCount = 0;
                }
                if ( !onlyDocument )
                {
                    try
                    {
                        bookDocuments[i].close( SaveOptions.YES, docNames[i] );
                        bookDocuments[i] = app.open( docNames[i], false );
                    }
                    catch ( err )
                    {
                        progresspanel.hide()
                        alert( err, '오류' )
                        return -1;
                    }
                }
            }
            returnValue = BookContents( bookDocuments, sourceFile, XMLToolStatus, XMLStructureStatus, myImageSize );
            if ( returnValue < 0 )
                return returnValue;
            if ( XMLToolStatus == 2 ) // Export XML
            {
                for ( var i = 0 ; i < bookDocumentsLength ; i++ )
                {
                    //Get_XmlId( bookDocuments[i] );
                    Set_DestinationID( bookDocuments[i] );
                    if ( ExportXML( bookDocuments[i] ) < 1 )
                        resultLog += '\r\n' + bookDocuments[i].name;
                    else
                    {
                        var delLog = new File( bookDocuments[i].filePath + '/Log/' + bookDocuments[i].name.split( '.' )[0] + '_Log.txt' );
                        delLog.remove();
                    }
                }
            }
            break;
        case 4: //Delete Tags
            returnValue = BookContents( bookDocuments, sourceFile, XMLToolStatus, XMLStructureStatus, myImageSize );
            if ( returnValue < 0 )
                return returnValue;
            break;
        default:
            break;
    }

    if ( onlyDocument ) {
        for ( var i = 0 ; i <app.documents.count() ; i++ ) {
            if ( !app.documents[i].visible ) {
                app.documents[i].close( SaveOptions.NO );
                i = -1;
            } else {
                try {
                    if (docPath.xpath("//*[@path = \'" + app.documents[i].fullName + "\']").@path.toString() != app.documents[i].fullName) {
                        app.documents[i].close( SaveOptions.NO );
                        i = -1;
                    }
                } catch (err) {
                    app.documents[i].close( SaveOptions.NO );
                    i = -1;
                }
            }
        }
    }
    else
    {
        for ( var i = 0 ; i < bookDocumentsLength ; i++ )
        {
            try { bookDocuments[i].close( SaveOptions.YES, docNames[i] ); }
            catch ( err )
            {
                progresspanel.hide();
                alert( docNames[i].name + '의 저장에 실패했습니다.\r\nerr' );
            }
        }
        for ( var i = 0 ; i < app.documents.count() ; i++ )
        {
            var item = app.activeBook.bookContents.itemByName(app.documents[i].name);
            if ( app.documents[i].fullName != item.fullName )
            {
                app.documents[i].close();
                i = -1;
            }
        }        
        app.activeBook.save();
    }
    progresspanel.hide();
    return 1;
}

function BookContents( bookDocuments, sourceFile, XMLToolStatus, XMLStructureStatus, myImageSize )
{
    var bookDocumentsLength = bookDocuments.length;
    var returnValue = 1;
    var value = null;
    for ( var i = 0 ; i < bookDocumentsLength ; i++ )
    {
        progresspanel.progressbar.value = i + 1;
        value= BookXML( bookDocuments[i], sourceFile, XMLToolStatus, XMLStructureStatus, myImageSize );
        if ( value < -1 )
        {
            progresspanel.hide();
            var docid = bookDocuments[i].id;
            alert( bookDocuments[i].name + ' 문서에서 스타일 에러가 발생했습니다' );
            for ( var j = 0; j < bookDocumentsLength ; j++ )
            {
                if ( bookDocuments[j].id != docid )
                    bookDocuments[j].close( SaveOptions.NO );
            }
            return -1;
        }
        returnValue = returnValue < value ? returnValue : value
    }
    return returnValue;
}

function Delete_Hyperlink( mydoc )
{
    var hyper_count = mydoc.hyperlinkURLDestinations.count();
    for ( var i = 0 ; i < hyper_count  ; i++ )
        mydoc.hyperlinkURLDestinations[0].remove();
    hyper_count = mydoc.crossReferenceSources.count();
    for ( var i = 0 ; i < hyper_count  ; i++ )
        mydoc.crossReferenceSources[0].remove();
    hyper_count = mydoc.hyperlinkTextDestinations.count();
    for ( var i = 0 ; i < hyper_count ; i++ )
        mydoc.hyperlinkTextDestinations[0].remove();
    hyper_count = mydoc.hyperlinks.count();
    for ( var i = 0 ; i < hyper_count ; i++ )
        mydoc.hyperlinks[0].remove();
}

//Book파일과 INDD파일을 선택해주는 함수를 호출하는 함수
function SelectTarget( allDocument )
{
    try
    {
        var books = app.books;
        var booksCount = app.books.count();
        if ( booksCount > 1 )
        {
            var book = SelectBook( books );
            if ( book == 0 )
                return 0;
            else if ( book < 0 )
                return book;
            app.activeBook = book;
        }
        else if ( booksCount == 1 )
            book = app.activeBook;
        else
        {
            alert( '열려진 북 파일이 없습니다', '오류' )
            return -1;
        }
        var bookContents = book.bookContents;
        var bookContentsCount = null;
        try { bookContentsCount = bookContents.count(); }
        catch ( err )
        {
            if ( !book )
                return 0;
            else
                return -1;
        }
        var docNames = [];
        if ( allDocument )
        {
            for ( var i = 0 ; i < bookContentsCount ; i++ )
                docNames[i] = bookContents[i].fullName;
        }
        else
        {
            for ( var i = 0 ; i < bookContentsCount ; i++ )
                docNames[i] = bookContents[i];
            docNames = SelectINDD( docNames );
            if ( !docNames ) return 0;
            else if ( docNames < 0 ) return -1;
        }
    } catch ( err )
    {
        alert( '열려있는 문서나 북이 없습니다', '오류' );
        return -1;
    }
    return docNames;
}
//Book파일이 여러개 일 경우 선택
function SelectBook( books )
{
    var value = 0;
    var res =
        "dialog { alignChildren: 'fill', \
            Alert: Panel { orientation: 'column', alignChildren:'left', \
                text: '다음 BOOK 중에 하나를 선택해주세요', \
                s: StaticText { text: '열번째 이후의 BOOOK은 표시되지 않습니다' },\
            }, \
            Books: Panel { orientation: 'column', alignChildren:'center', \
                text: '문서 종류', \
            }, \
            buttons: Group { orientation: 'row', alignment: 'right', \
                backBtn: Button { text:'Back', alignment:'left' }, \
                cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                okBtn: Button { text:'Next', properties:{name:'ok'} }, \
            } \
        }";
    var select = new Window( res, 'BOOK 파일이 여러개 열려있습니다' );
    select.hepTip = '목록의 BOOK 중에 하나를 선택합니다\r\n10개 이상의 BOOK이 동시에 열려 있을 경우 10개까지만 표시를 합니다\r\n파일 중 하나를 선택하면 다음 단계로 넘어갑니다';
    var booksCount = books.count();
    var myBook = [];
    for ( var i = 0 ; i < booksCount ; i++ )
    {
        myBook[i] = select.children[0].add( 'button', undefined, books[i].name.split( '.' )[0] );
        myBook[i].minimumSize = [222, 25]
    };
    select.buttons.backBtn.onClick = function () { select.close( -1 ) };
    select.buttons.cancelBtn.active = false;
    select.buttons.cancelBtn.visible = false;
    select.buttons.ok.active = false;
    select.buttons.ok.active = false;
    select.center();
    try
    {
        myBook[0].onClick = function () { value = 0; select.close( 1 ); }
        myBook[1].onClick = function () { value = 1; select.close( 1 ); }
        myBook[2].onClick = function () { value = 2; select.close( 1 ); }
        myBook[3].onClick = function () { value = 3; select.close( 1 ); }
        myBook[4].onClick = function () { value = 4; select.close( 1 ); }
        myBook[5].onClick = function () { value = 5; select.close( 1 ); }
        myBook[6].onClick = function () { value = 6; select.close( 1 ); }
        myBook[7].onClick = function () { value = 7; select.close( 1 ); }
        myBook[8].onClick = function () { value = 8; select.close( 1 ); }
        myBook[9].onClick = function () { value = 9; select.close( 1 ); }
    } catch ( err ) { }
    var result = select.show() - 1;
    if ( result < 0 )
        return 0;
    else if ( result > 0 )
        return -1;
    return books[value];
}

//Book파일로 작업시, 필요한 INDD파일 선택
function SelectINDD( INDD )
{
    var value = 0;
    var res =
        "dialog { alignChildren: 'fill', \
            Alert: Panel { orientation: 'column', alignChildren:'left', \
                text: '다음 INDD 중 필요한 파일을 선택해 주세요', \
            }, \
            INDD: Panel { orientation: 'column', alignChildren:'left', \
                text: 'INDD 파일', \
            }, \
            SelectAll: Panel { orientation: 'column', \
                text: 'Select All or Deselect All',\
                Select: Group { orientation: 'row', \
                    All: RadioButton { text: 'Select All' },\
                    DeAll: RadioButton { text: 'Deselect All' },\
                }\
            }, \
            buttons: Group { orientation: 'row', alignment: 'right', \
                backBtn: Button { text:'Back'}, \
                cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                okBtn: Button { text:'Next', properties:{name:'ok'} }, \
            } \
        }";
    var select = new Window( res, 'INDD파일 선택창' );
    select.helpTip = '필요한 인디자인 파일을 선택하면 다음 단계로 넘어갑니다';
    var INDDCount = INDD.length;
    var myINDD = [];
    for ( var i = 0 ; i < INDDCount ; i++ )
    {
        myINDDName = INDD[i].fullName.name.split( '%20' ).join( ' ' );
        myINDD[i] = select.INDD.add( 'checkbox', undefined, myINDDName.split( '.' )[0] );
        if ( ( myINDDName.split( '.' )[0] != 'TOC' ) && ( myINDDName.split( '.' )[0] != 'IX' ) )//&& (myINDDName.split('.')[0].search('TTE')<0) && (myINDDName.split('.')[0].search('Safety')<0) )
            myINDD[i].value = true;
        //myINDD[i].minimumSize = [120, 30];
    }
    var myINDDLength = myINDD.length;
    select.SelectAll.Select.All.onClick = function ()
    {
        for ( var i = 0 ; i < myINDDLength ; i++ )
            myINDD[i].value = true;
    }
    select.SelectAll.Select.DeAll.onClick = function ()
    {
        for ( var i = 0 ; i < myINDDLength ; i++ )
            myINDD[i].value = false;
    }
    select.buttons.backBtn.onClick = function () { select.close( -1 ) };
    select.buttons.cancelBtn.active = false;
    select.buttons.cancelBtn.visible = false;
    select.center();
    myActiveINDD = [];
    //var result = select.show()-1;
    //if (result>0) return -1;
    //else if (result<0) return 0;
    for ( var i = 0 ; i < myINDDLength ; i++ )
    {
        if ( myINDD[i].value )
            myActiveINDD.push( INDD[i].fullName );
    }
    return myActiveINDD;
}

//XML호출시 호출할 영문 INDD파일을 선택
function SelectSourceINDD( myDoc )
{
    var myDocName = myDoc.name.split( '%20' ).join( ' ' ).split( '.' )[0];
    var sourceFile = new Window( 'dialog', '영문 인디자인 파일을 선택해주세요' );
    sourceFile.center();
    sourceFile = File.openDialog( 'Select ' + myDocName + ' File for Indexing', 'InDesign Files:' + myDocName + '.indd' );
    if ( sourceFile )
        return sourceFile;
    else return 0;
}

//색인 작업의 여부를 결정
function IndexFunctionSelect()
{
    var res =
        "dialog { alignChildren: 'fill', \
            Work: Panel { orientation: 'column', alignChildren:'center', \
                text: 'XML 호출이 끝난 후 인덱스 작업을 하시겠습니까?', alignment: 'center', \
                OK: Button { text:'인덱스 작업을 실행합니다', minimumSize:[200, 35] }, \
                CANCER: Button { text:'인덱스 작업을 하지 않습니다', minimumSize:[200, 35] },\
            }, \
        }";
    select = new Window( res, '색인 작업 선택 창' );
    select.Work.OK.onClick = function () { IndexCondition = 1; select.close( 4 ) };
    select.Work.CANCER.onClick = function () { select.close( 3 ) };
    select.helpTip = '색인 작업 여부를 결정합니다\r\n취소를 누를 경우 XML 호출 작업만 진행합니다';
    select.Work.OK.helpTip = 'XML 호출이 끝난 후, 해당 문서의 영문 버전 Indd 파일을 선택해서 색인 정보를 불러옵니다\r\n북에서 작업하는 경우, 한번 선택하면 이후는 자동으로 선택이 됩니다\r\n색인 정보가 없는 경우 XML 호출만 합니다';
    select.Work.OK.helpTip = 'XML 호출만 수행합니다. 작업이 끝난 후, 태그와 구조가 지워지기 때문에 이후 색인 작업을 하기 위해서는 Side Tools 를 이용해야 합니다';
    select.center();
    var result = select.show() - 3;
    if ( result < 0 )
        return -1;
    else if ( result )
        return 1;
    else
        return 0;
}

// XML 호출, 생성, 삭제 기능 중 하나를 선택
function SelectXMLTools()
{
    var XMLToolStatus = null;
    var res = null;
    var result = null;
    var type = null;
    if ( !integrate )
    {
        var res =
            "dialog { alignChildren: 'fill', \
                Main: Panel { alignChildren:'center', \
                    text: '편집 메뉴', alignment: 'center', \
                    Option1: Group { orientation: 'row', alignment: 'center', \
                        ImportXML: Button { text: 'Import XML', minimumSize:[125, 25] }, \
                        ExportXML: Button { text: 'Export XML', minimumSize:[125, 25] },\
                    } \
                    Option2: Group { orientation: 'row', alignment: 'center', \
                        CheckXML: Button { text: 'Error Check', minimumSize:[125, 25] },\
                        DelXML: Button { text: 'Delete Tag', minimumSize:[125, 25] },\
                    } \
                }, \
                Sub: Panel { alignChildren:'center', \
                    text: 'HTML 메뉴', alignment: 'center', \
                    Option1: Group { orientation: 'row', alignment: 'center', \
                        Select: Button { text: 'Make C_Image', minimumSize: [125, 25] }, \
                        DeSelect: Button { text: 'Delete C_Image', minimumSize: [125, 25] }, \
                    } \
                    Option2: Group { orientation: 'row', alignment: 'center', \
                        ReDefine: Button { text: 'Redefine Styles', minimumSize: [125, 25] }, \
                        ExportIDML: Button { text: 'Export IDML/INX', minimumSize:[125, 25] },\
                    } \
                }, \
                buttons: Group { orientation: 'row', \
                    backBtn: Button { text:'뒤로', alignment: 'left' }, \
                    ExportAll: Checkbox { text: 'cover, TOC 포함 (전체 IDML)'},\
                } \
            }";
        var select = new Window( res, '일반 메뉴얼' );
        select.center();
        select.Main.Option1.ImportXML.onClick = function () { type = 1; select.close( 1 ) };
        select.Main.Option1.ExportXML.onClick = function () { type = 2; select.close( 1 ) };
        select.Main.Option2.CheckXML.onClick = function () { type = 3; select.close( 1 ) };
        select.Main.Option2.DelXML.onClick = function () { type = 4; select.close( 1 ) };
        select.Sub.Option1.Select.onClick = function () { type = 5; select.close( 1 ) };
        select.Sub.Option1.DeSelect.onClick = function () { type = 6; select.close( 1 ) };
        select.Sub.Option2.ReDefine.onClick = function () { type = 7; select.close( 1 ) };
        select.Sub.Option2.ExportIDML.onClick = function () { type = 8; select.close( 1 ) };
        select.buttons.backBtn.onClick = function () { select.close( -1 ) };
        select.helpTip = '필요한 기능을 클릭하면 다음 단계로 넘어갑니다';
        select.Main.Option1.ImportXML.helpTip = '문서가 있는 폴더에서 XML 폴더 안의\r\n'
                                                                    + '같은 이름을 가진 xml파일을 호출하여\r\n'
                                                                    + '다국어 작업을 해주는 기능입니다\r\n'
                                                                    + '색인 작업을 할수 있습니다';
        select.Main.Option1.ExportXML.helpTip = '문서가 있는 풀더에서 XML 폴더를 만들고\r\n같은 이름의 xml파일을 만들어 줍니다';
        select.Main.Option2.CheckXML.helpTip = '알려진 오류가 있는지 체크한 뒤 태그를 지워줍니다';
        select.Main.Option2.DelXML.helpTip = '태그를 지워줍니다';
        select.Sub.Option1.Select.helpTip = 'C_Image 스타일을 생성합니다\r\nHTML관련 작업이 필요할때 실행해주세요'
        select.Sub.Option1.DeSelect.helpTip = 'C_Image 스타일을 지워줍니다\r\nC_Image와 관련된 오류가 나타날 경우 실행해주세요'
        select.Sub.Option2.ReDefine.helpTip = '설정된 GREP\r\n'
                                                                + 'NoteToCellBody'
                                                                + 'Refer_CharStyle_del\r\n'
                                                                + 'Refer_CharStyle_del2\r\n'
                                                                + 'del_mmi_nextcolumn\r\n'
                                                                + 'del_mmi_nobold_nextcolumn\r\n';

        select.Sub.Option2.ExportIDML.helpTip = 'INDD파일을 IDML파일로 변환할 때 쓰입니다'
        result = select.show() - 1;
        export_all = select.buttons.ExportAll.value;
    }

    if ( result == 0 )
        return type;
    else if ( result > 0 )
        return -1;
    else
        return 0;
}

//XML을 생성할 시, 일반적인 형태로 만들 것인지, 단계적인 형태로 만들 것인지 결정
function SelectXMLStructure()
{
    var res =
        "dialog { alignChildren: 'fill', \
            Type: Panel { orientation: 'column', alignChildren:'center', \
                text: 'XML 구조',\
                SelectType: Group { orientation: 'row', alignChildren:'center', \
                    Nomal: Button { text: 'Nomal', minimumSize: [80, 25] },\
                    Tree: Button { text: 'Tree', minimumSize: [80, 25] },\
                }\
            }, \
            CheckMMI: Panel { orientation: 'column', alignChildren:'left', \
                MMI: Checkbox { text: 'MMI Count'},\
                Trim: Checkbox { text: '문자 스타일 적용 텍스트의 앞 뒤 공백 스타일 제거'},\
            },\
            buttons: Group { orientation: 'row', \
                backBtn: Button { text:'뒤로', alignment: 'left' }, \
                doNotCheckError : Checkbox { text:'스타일 오류 확인하지 않기', aligment: 'right'}, \
            }\
        }";
    var style;
    var select = new Window( res, 'XML 구조화' );
    select.center();
    select.helpTip = '구조를 일반적인 형태로 만들지, 단계화 할지 정합니다';
    select.Type.SelectType.Nomal.helpTip = '일반적인 형태의 구조를 만듭니다\r\n스타일에 태그를 입힌 후에 구조를 만들어줍니다\r\n작업이 끝난 후 XML을 추출, 작성합니다';
    select.Type.SelectType.Tree.helpTip = '일반적인 형태의 구조를 만든 후에\r\n같은 그룹끼리 묶어주며 단계화를 해줍니다\r\n작업이 끝난 후 XML을 추출, 작성합니다';
    select.Type.SelectType.Nomal.onClick = function ()
    {
        if ( select.CheckMMI.MMI.value )
            style = 2;
        else
            style = 1;
        select.close( 1 );
    }
    select.Type.SelectType.Tree.onClick = function ()
    {
        if ( select.CheckMMI.MMI.value )
            style = 4;
        else
            style = 3;
        select.close( 1 );
    }
    select.buttons.backBtn.onClick = function () { select.close( -1 ) };
    var result = select.show() - 1;
    doNotCheckErrors = select.buttons.doNotCheckError.value;
    trimCharacterStyleContents = select.CheckMMI.Trim.value;
    if ( !result )
        return style;
    else if ( result > 0 )
        return -1;
    else
        return 0;
}

//Book파일을 대상으로 XML작업
function BookXML( myDoc, sourceFile, XMLToolStatus, XMLStructureStatus, myImageSize )
{
    var reVal = null;
    var checkExport = 0;
    var checkImport = 0;
    var log = null;
    var copyXML = {};
    Temp_XML = new Array();    
    for ( var i = 0; i < myDoc.xmlElements[0].xmlAttributes.count() ; i++ )
    {
        copyXML = new Object();
        copyXML.name = myDoc.xmlElements[0].xmlAttributes[i].name.toString();
        copyXML.value = myDoc.xmlElements[0].xmlAttributes[i].value.toString();
        if ( copyXML.name.search( 'Cell_' ) > -1 )
            Temp_XML = Temp_XML.concat( copyXML );
    }
    switch ( XMLToolStatus )
    {
        case 1:
            if (sourceFile != 0)
                sourceFile = File( sourceFile.path + '/' + myDoc.name );
            checkImport = ImportXML( myDoc, sourceFile, myImageSize );
            Set_HyperlinkDestination(myDoc);
            SetPageBreak(myDoc);
            if ( checkImport < 0 )
            {
                if ( errorCount )
                {
                    resultLog += '\r\n' + myDoc.name;
                    errorCount = 0;
                    totalError++;
                }
                else
                    totalError++;
                reVal = -1;
            }
            else
                reVal = 1;
            break;
        case 2:
        case 3:
            if ( XMLStructureStatus < 3 )
                checkExport = NomalStructure( myDoc );
            else
                checkExport = TreeStructure( myDoc );
            if ( errorCount )
            {
                if ( styleErrorCount )
                {
                    alert( '잘못된 스타일이 발견되어 프로그램을 종료합니다', '오류' );
                    reVal = -2; // return -2; }
                }
                else
                {
                    resultLog += '\r\n' + myDoc.name;
                    errorCount = 0;
                    totalError++;
                    reVal = -1;
                }
            }
            else if ( XMLToolStatus == 3 )
            {  // Check XML
                DeleteTag( myDoc );    // 태그를 지워준다
                reVal = 3;
            }
            else
            {  // Export XML
                if ( checkExport )
                {
                    copyXML = new Object();
                    var date = new Date();
                    copyXML.name = "ExportDate";
                    copyXML.value = date.toString();
                    Temp_XML = Temp_XML.concat( copyXML );
                    if (doNotCheckErrors) {
                        var doNotErrorChecking = new Object();
                        doNotErrorChecking.name = "doNotErrorChecking";
                        doNotErrorChecking.value = "true";
                        Temp_XML = Temp_XML.concat(doNotErrorChecking);
                    }
                    for ( var i = 0; i < Temp_XML.length ; i++ )
                        myDoc.xmlElements[0].xmlAttributes.add( Temp_XML[i].name, Temp_XML[i].value );

                    return 2;
                }
                reVal = 2;
            }
            break;
        case 4:                 // 태그를 지울 경우 실행
            DeleteTag( myDoc );   // 태그를 지워준다
            reVal = 4;
            break;
        default:
            reVal = -1;
            break;
    }
    try
    {
        for ( var i = 0; i < Temp_XML.length ; i++ )
            myDoc.xmlElements[0].xmlAttributes.add( Temp_XML[i].name, Temp_XML[i].value )
    }
    catch ( err ) { }
    return reVal;
}

// 텍스트 앵커 새로이 생성
function Set_HyperlinkDestination(mydoc)
{
    var hyper_count = mydoc.hyperlinkTextDestinations.count();
    for ( var i = 0 ; i < hyper_count ; i++ )
        mydoc.hyperlinkTextDestinations[0].remove();
    var anchors = mydoc.xmlElements[0].evaluateXPathExpression( '//*[@hyperlinkInform = \'TextAnchor\']' );
    var anchor_count = anchors.length;
    for ( var i = 0; i < anchor_count ; i++ )
    {
        var anchor = anchors[i];
        var item = mydoc.hyperlinkTextDestinations.add(anchor.texts[0]);
        item.name = anchor.xmlAttributes.itemByName( 'name' ).value;
        item.label = anchor.xmlAttributes.itemByName( 'label' ).value + '';
    }
    return 1;
}

//페이지 나누기 설정
function SetPageBreak(mydoc)
{
    var root = mydoc.xmlElements.item( 0 );
    var pbs = root.evaluateXPathExpression( '//*[@paragraphBreakType = \'NextColumn\']' );
    //grep_setting("\\r", null, null, "~M\\r");
    for (var i = 0 ; i < pbs.length ; i++) 
    {
        pbs[i].insertTextAsContent(SpecialCharacters.COLUMN_BREAK, XMLElementPosition.ELEMENT_END );
        if (pbs[i].characters[pbs[i].characters.length - 2].contents == '\r')
            pbs[i].characters[pbs[i].characters.length - 2].remove();
        //pbs[i] .changeGrep();
        //pbs[i] .paragraphs[pbs[i] .paragraphs.length - 1].remove();
    }
}

function Save_Destination(mydoc)
{
    var docs = [];
    linked = [];
    var files = new Folder(mydoc.filePath).getFiles('*.indd');
    for ( var i = 0 ; i < files.length ; i++)
    {
        if (files[i].toString() != mydoc.fullName)
            docs.push(app.open(files[i], false));
    }
    for ( var i = 0 ; i < docs.length ; i++)
    {
        var item = docs[i];
        for ( var j = 0 ; j < item.hyperlinks.count() ; j++)
        {
            try
            {
                if ( item.hyperlinks[j].source.constructor == CrossReferenceSource && item.hyperlinks[j].destination.parent.id == mydoc.id)
                    linked.push([item.id, item.hyperlinks[j], item.hyperlinks[j].destination.name]);
            }
            catch (err)
            {
                alert (err.toString() + '\r\n오류가 난 문서 : ' + item.name + '\r\n오류가 난 링크 : ' + item.hyperlinks[j], '링크 오류');
            }
        }
    }
}

function Reconnect_Destination(mydoc)
{    
    for ( var i = 0 ; i < linked.length ; i++)
    {
        linked[i][1].destination = mydoc.hyperlinkTextDestinations.itemByName(linked[i][2].toString());
        app.documents.itemByID(linked[i][0] * 1).save();
    }
    for ( var i = 0 ; i < app.documents.count() ; i++)
    {
        if (!app.documents[i].visible)
        {
            app.documents[i].close(SaveOptions.NO);
            i = -1;
        }
    }
}

// XML본문 끝-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function mainSide()
{
    var res =
        "dialog { alignChildren: 'fill', \
            Option: Panel { alignChildren:'center', \
                text: '메뉴', \
                Option1: Group { orientation: 'row', alignment: 'center', \
                    FindErrorPoint: Button { text: 'Find Error Point', minimumSize: [155, 25] }, \
                    ImportIndex: Button { text: 'Make Index', minimumSize: [155, 25] },\
                } \
                Option2: Group { orientation: 'row', alignChildren:'center', \
                    ResizeImage: Button { text: 'Image Resize', minimumSize: [155, 25] },\
                    RelinkImage: Button { text: 'Image Link', minimumSize: [155, 25] },\
                } \
                Option3: Group { orientation: 'row', alignChildren:'center', \
                    CheckH: Button { text: 'Check Highlight', minimumSize: [155, 25] },\
                    UncheckH: Button { text: 'Uncheck Highlight', minimumSize: [155, 25] },\
                } \
        		Option5: Group { orientation: 'row', alignChildren:'left', \
                    CheckPara: Button { text: 'Check Highlight-Eng', minimumSize: [155, 25] },\
                    UncheckPara: Button { text: 'Uncheck Highlight-Eng', minimumSize: [155, 25] },\
                } \
                Option4: Group { orientation: 'row', alignChildren:'left', \
                    UncheckMMI: Button { text: 'Uncheck Changed MMI', minimumSize: [155, 25] },\
                    CheckMMI: Button { text: 'Check Changed MMI', minimumSize: [155, 25] },\
                } \
            }, \
            buttons: Group { orientation: 'row', alignment: 'left', \
                backBtn: Button { text:'뒤로', alignment:'left' }, \
            } \
        }";
    var selectWindow = new Window( res, '기능 선택 창' );
    selectWindow.buttons.backBtn.onClick = function () { selectWindow.close( -1 ) };
    selectWindow.Option.Option1.FindErrorPoint.onClick = function () { SelectStatus = 1; selectWindow.close( 1 ) };
    selectWindow.Option.Option1.ImportIndex.onClick = function () { SelectStatus = 2; selectWindow.close( 1 ) };
    selectWindow.Option.Option2.ResizeImage.onClick = function () { SelectStatus = 3; selectWindow.close( 1 ) };
    selectWindow.Option.Option2.RelinkImage.onClick = function () { SelectStatus = 4; selectWindow.close( 1 ) };
    selectWindow.Option.Option3.CheckH.onClick = function () { SelectStatus = 5; selectWindow.close( 1 ) };
    selectWindow.Option.Option3.UncheckH.onClick = function () { SelectStatus = 6; selectWindow.close( 1 ) };
    selectWindow.Option.Option4.UncheckMMI.onClick = function () { SelectStatus = 7; selectWindow.close( 1 ) };
    selectWindow.Option.Option4.CheckMMI.onClick = function () { SelectStatus = 8; selectWindow.close( 1 ) };
    selectWindow.Option.Option5.CheckPara.onClick = function() { SelectStatus = 9; selectWindow.close(1) };
    selectWindow.Option.Option5.UncheckPara.onClick = function() { SelectStatus = 10; selectWindow.close(1) };
    selectWindow.Option.Option4.CheckMMI.visible = false;
    selectWindow.helpTip = '필요한 기능을 클릭하면 다음 단계로 넘어갑니다';
    selectWindow.Option.Option1.FindErrorPoint.helpTip = '로그 파일을 이용해 오류가 난 곳을 찾습니다\r\n로그 파일은 해당 문서의 Log 폴더에 있습니다';
    selectWindow.Option.Option1.ImportIndex.helpTip = '색인 작업을 수행합니다\r\nXML을 추출하거나 호출하지 않습니다\r\n반드시 UM만 작업해야 합니다';
    selectWindow.Option.Option2.ResizeImage.helpTip = '이미지들의 크기를 설정해줍니다';
    selectWindow.Option.Option2.RelinkImage.helpTip = '이미지들의 링크를 재설정 해줍니다';
    selectWindow.center();

    var select = selectWindow.show() - 1;
    if ( !select )
    {
        switch ( SelectStatus )
        {
            case 1:
                try
                {
                    FindErrorPoint();
                    break;
                }
                catch ( err )
                {
                    alert( '열어놓은 문서가 없습니다', '문서를 찾을 수 없습니다' );
                    return -1;
                }
            case 2:
                alert( 'UM 문서가 아닐 경우 오류가 발생합니다', 'UM 확인창 ' );
                return ImportIndex();
            case 3:
                var myDoc = null;
                if ( app.documents.count() > 0 )
                {
                    try
                    {
                        myDoc = app.activeDocument;
                        myDoc.name;
                    }
                    catch ( err )
                    {
                        try
                        {
                            myDoc = app.documents[0];
                            myDoc.name;
                        }
                        catch ( err )
                        {
                            alert( '열어놓은 문서가 없습니다', '문서를 찾을 수 없습니다' );
                            return -1;
                        }
                    }
                    try { ResizeImage( myDoc, +InputImageSize() ); }
                    catch ( err )
                    {
                        alert( '열어놓은 문서가 없습니다', '문서를 찾을 수 없습니다' );
                        return -1;
                    }
                }
                else if ( app.books.count() > 0 )
                {
                    var mybook = null;
                    try
                    {
                        mybook = app.activeBook;
                    }
                    catch ( err )
                    {
                        mybook = app.books[0];
                    }
                    var bookcontents = mybook.bookContents;
                    var contentscount = bookcontents.count();
                    var imagesize = InputImageSize();
                    for ( var i = 0 ; i < contentscount ; i++ )
                    {
                        myDoc = app.open( bookcontents[i].fullName, false );
                        ResizeImage( myDoc, +imagesize );
                        myDoc.close( SaveOptions.YES );
                    }
                }
                else
                {
                    alert( '열어놓은 북이나 문서가 없습니다', '문서를 찾을 수 없습니다' );
                    return -1;
                }
                alert( '완료되었습니다', '크기 조정' )
                break;
            case 4:
                return RelinkImage(); break;
            case 5:
                HighlightChecker( true );
                break;
            case 6:
                HighlightChecker( false );
                break;
            case 7:
                Check_Changed_MMI( false );
                break;
            case 8:
                Check_Changed_MMI( true );
                break;
            case 9:
            	Paragraph_HighlightChecker(true);
            	break;
            case 10:
            	Paragraph_HighlightChecker(false);
            	break;
        }
    }
    else if ( select < 0 )
        return 0;
    return 1;
}
// SIDE본문 끝-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function EditIntegrate()
{
    var condition;
    do
    {
        integrate = true;
        condition = 1;
        var res =
            "dialog { alignChildren: 'fill', \
                Xml: Panel { alignChildren:'center', \
                    text: '편집 메뉴', \
                    Group1: Group { orientation: 'row', alignment: 'center', \
                        ImportXML: Button { text: 'Import XML', minimumSize:[125, 25] }, \
                        ExportXML: Button { text: 'Export XML', minimumSize:[125, 25] },\
                    } \
                    Group2: Group { orientation: 'row', alignment: 'center', \
                        CheckXML: Button { text: 'Error Check', minimumSize:[125, 25] },\
                        DelXML: Button { text: 'Delete Tag', minimumSize:[125, 25] },\
                    } \
                }, \
                Option: Panel { alignChildren:'center', \
                    text: 'HTML 메뉴', \
                    Group1: Group { orientation: 'row', alignment: 'center', \
                        Select: Button { text: 'Make C_Image', minimumSize: [125, 25] }, \
                        DeSelect: Button { text: 'Delete C_Image', minimumSize: [125, 25] }, \
                    } \
                    Group2: Group { orientation: 'row', alignment: 'center', \
                        ReDefine: Button { text: 'Redefine Styles', minimumSize: [125, 25] }, \
                        ExportIDML: Button { text: 'Export IDML/INX', minimumSize: [125, 25] }, \
                    } \
                    Group3: Group { orientation: 'row', alignment: 'center', \
                        MMItoI: Button { text: 'MMI -> Iword', minimumSize: [125, 25] }, \
                        ItoMMI: Button { text: 'Iword -> MMI', minimumSize: [125, 25] }, \
                    } \
                }, \
                buttons: Group { orientation: 'row', alignment: 'left', \
                    backBtn: Button { text:'뒤로', alignment:'left' }, \
                    QSG: Checkbox { text: 'QSG'},\
                    ExportAll: Checkbox { text: 'cover, TOC 포함 (전체 IDML)'},\
                } \
            }";

        var selectWindow = new Window( res, '통합 메뉴얼' );
        selectWindow.buttons.backBtn.onClick = function () { selectWindow.close( -1 ) };
        selectWindow.Xml.Group1.ImportXML.onClick = function () { SelectStatus = 1; selectWindow.close( 1 ) };
        selectWindow.Xml.Group1.ExportXML.onClick = function () { SelectStatus = 2; selectWindow.close( 1 ) };
        selectWindow.Xml.Group2.CheckXML.onClick = function () { SelectStatus = 3; selectWindow.close( 1 ) };
        selectWindow.Xml.Group2.DelXML.onClick = function () { SelectStatus = 4; selectWindow.close( 1 ) };
        selectWindow.Option.Group1.Select.onClick = function () { SelectStatus = 5; selectWindow.close( 1 ) };
        selectWindow.Option.Group1.DeSelect.onClick = function () { SelectStatus = 6; selectWindow.close( 1 ) };
        selectWindow.Option.Group2.ReDefine.onClick = function () { SelectStatus = 7; selectWindow.close( 1 ) };
        selectWindow.Option.Group2.ExportIDML.onClick = function () { SelectStatus = 8; selectWindow.close( 1 ) };
        selectWindow.Option.Group3.MMItoI.onClick = function () { SelectStatus = 9; selectWindow.close( 1 ) };
        selectWindow.Option.Group3.ItoMMI.onClick = function () { SelectStatus = 10; selectWindow.close( 1 ) };

        selectWindow.Xml.Group1.ImportXML.helpTip = '문서가 있는 폴더에서 XML 폴더 안의\r\n'
                                                                    + '같은 이름을 가진 xml파일을 호출하여\r\n'
                                                                    + '다국어 작업을 해주는 기능입니다\r\n'
                                                                    + '색인 작업을 할수 있습니다';
        selectWindow.Xml.Group1.ExportXML.helpTip = '문서가 있는 풀더에서 XML 폴더를 만들고\r\n같은 이름의 xml파일을 만들어 줍니다';
        selectWindow.Xml.Group2.CheckXML.helpTip = '알려진 오류가 있는지 체크한 뒤 태그를 지워줍니다';
        selectWindow.Xml.Group2.DelXML.helpTip = '태그를 지워줍니다';
        selectWindow.Option.Group1.Select.helpTip = 'C_Image 스타일을 생성합니다\r\nHTML관련 작업이 필요할때 실행해주세요'
        selectWindow.Option.Group1.DeSelect.helpTip = 'C_Image 스타일을 지워줍니다\r\nC_Image와 관련된 오류가 나타날 경우 실행해주세요'
        selectWindow.Option.Group2.ReDefine.helpTip = '설정된 GREP\r\n'
                                                                                + 'C_image\r\n'
                                                                                + 'C_image_2\r\n'
                                                                                + 'C_image_3\r\n'
                                                                                + 'C_image_4\r\n'
                                                                                + 'C_image_5\r\n'
                                                                                + 'Refer_CharStyle_del\r\n'
                                                                                + 'Refer_CharStyle_del2\r\n'
                                                                                + 'del_mmi_nextcolumn\r\n'
                                                                                + 'del_mmi_nobold_nextcolumn\r\n'
                                                                                + 'del_blank_after_sentence_Common\r\n'
                                                                                + 'del_blank_after_table1_Common\r\n'
                                                                                + 'del_blank_after_table2_Common';
        selectWindow.Option.Group2.ExportIDML.helpTip = 'INDD파일을 IDML파일로 변환할 때 쓰입니다'
        selectWindow.Option.Group3.MMItoI.helpTip = 'MMI 스타일 이름을 Iword 로 변경합니다'
        selectWindow.Option.Group3.ItoMMI.helpTip = 'Iword 스타일 이름을 MMI 로 변경합니다'
        selectWindow.center();

        var select = selectWindow.show() - 1;
        if ( !select )
        {
            qsg_type = selectWindow.buttons.QSG.value;
            export_all = selectWindow.buttons.ExportAll.value;
            no_intergrate = false;
            switch ( SelectStatus )
            {
                case 1:
                    condition = EditXML( 1 );
                    break;
                case 2:
                    condition = EditXML( 2 );
                    break;
                case 3:
                    condition = EditXML( 3 );
                    break;
                case 4:
                    condition = EditXML( 4 );
                    break;
                case 5:
                    try { alert( Make_Image(), 'C_Image 생성 결과' ); }
                    catch ( err ) { alert( '열어놓은 문서가 없습니다', '문서를 찾을 수 없습니다' ); return -1; }
                    break;
                case 6:
                    try { alert( Delete_Image(), 'C_Image 삭제 결과' ); }
                    catch ( err ) { alert( '열어놓은 문서가 없습니다', '문서를 찾을 수 없습니다' ); return -1; }
                    break;
                case 7:
                    try
                    {
                        var my_Queries = [
                            ["grep", "C_image", myNone],
                            ["grep", "C_image_2", myNone],
                            ["grep", "C_image_3", "MMI"],
                            ["grep", "C_image_4", "C_NoBreak"],
                            ["grep", "C_image_5", myNone],
                            ["grep", "Refer_CharStyle_del", "MMI_NoBold"],
                            ["grep", "Refer_CharStyle_del2", "MMI"],
                            ["grep", "Refer_CharStyle_del8", "C_NoBreak"],
                            ["grep", "del_mmi_nextcolumn", "MMI"],
                            ["grep", "del_mmi_nobold_nextcolumn", "MMI_NoBold"],
                            ["grep", "del_blank_Description-Cell", "Description-Cell"],
                            ["grep", "del_blank_Unorderlist_1-Cell", "UnorderList_1-Cell"]
                        ];
                        alert( Redefine_Style( my_Queries ), '참조 스타일 삭제 결과' );
                    } catch ( err ) { alert( '열어놓은 문서가 없습니다', '문서를 찾을 수 없습니다' ); return -1; }
                    break;
                case 8:
                    ExportIDML();
                    break;
                case 9:
                    ChangeInterface( 'MMI', 'Iword' );
                    break;
                default:
                    ChangeInterface( 'Iword', 'MMI' );
                    break;
            }
        }
        else if ( select < 0 ) { integrate = false; return 0; }
    } while ( !condition );
    integrate = false;
    return 1;
}

//XML에 오류가 있는 경우, XML구조에서 그 부분을 선택
function ShowErrors()
{
    var myDoc = null;
    try { myDoc = app.activeDocument; }
    catch ( err ) { return alert( '열어놓은 문서가 없습니다', '오류' ); }
    var errorAddress = [];
    var errorAddress2 = 0;
    var myRoot = myDoc.xmlElements.item( 0 );
    var myName = myDoc.name.toString().split( '.' );
    try { var FileInput = new File( myDoc.filePath + '/Log/' + myName[0] + '_Log' + '.txt' ); }
    catch ( err ) { return alert( '문서가 한 번도 저장되지 않았습니다.\r\n저장하신 후에 진행해주세요', '마침' ); }
    if ( FileInput.exists )
    {
        FileInput.open( "r" );
        myLog = FileInput.read().split( '\n' );
        FileInput.close();
        errorAddress = FindAddress( myLog );
        var errorCount = errorAddress.length;
        if ( !errorCount )
            return alert( '오류 주소가 없습니다', '마침' );
        else
        {
            errorAddress2 = errorAddress[0].split( ',' ).reverse();
            errorCount2 = errorAddress2.length - 1;
            errorPoint = ErrorSelect( myRoot, errorAddress2, errorCount2 );
            if ( errorPoint < 0 ) return -1;
            errorPoint.select();
            for ( var i = 1 ; i < errorCount ; i++ )
            {
                errorAddress2 = errorAddress[i].split( ',' ).reverse();
                errorCount2 = errorAddress2.length - 1;
                errorPoint = ErrorSelect( myRoot, errorAddress2, errorCount2 );
                if ( errorPoint < 0 ) return -1;
                errorPoint.select( SelectionOptions.ADD_TO );
            }
        }
        return 1;
    }
    else
    {
        alert( '오류 주소 파일이 없습니다', '마침' );
        return -1;
    }
}

function FindAddress( myLog )
{
    var newLog = [];
    var myLogCount = myLog.length
    for ( i = 0 ; i < myLogCount ; i++ )
        if ( myLog[i].search( 'Address' ) > -1 )
            newLog.push( myLog[i].split( ' ' )[1] );
    return newLog;
}

function ErrorSelect( mySection, errorAddress2, errorCount2 )
{
    if ( errorCount2 > -1 )
    {
        try { var mySection = mySection.xmlElements[errorAddress2[errorCount2]]; }
        catch ( err ) { alert( '오류 주소가 맞지가 않습니다. XML Export를 실행해서 로그를 업데이트 해주세요', '오류 주소가 틀림' ); return -1; }
        errorCount2--;
        return ErrorSelect( mySection, errorAddress2, errorCount2 );
    }
    else
        return mySection;
}

// ShowErrorPoint본문 끝-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

function Get_HyperlinkInform( mydoc )
{
    var link_count = mydoc.hyperlinks.count();
    var link_item = null;
    var source_item = null;
    var destination_item = null;
    var newitem = null;
    var stringtype = '';
    for ( var i = 0 ; i < link_count ; i++ )
    {
        link_item = mydoc.hyperlinks[i];
        source_item = link_item.source;
        destination_item = link_item.destination;
        
        if (source_item == null)
        {
            alert(mydoc.name + ' 파일에 잘못된 하이퍼 링크가 있습니다.');
        }
        else if (destination_item == null)
        {
            alert(mydoc.name + ' 파일에 잘못된 하이퍼 링크가 있습니다.');
        }
        else if ( source_item.constructor == CrossReferenceSource )
        {
            newitem = source_item.sourceText.associatedXMLElements[0].xmlElements.add( 'xref', source_item.sourceText );
            newitem.xmlAttributes.add( 'hyperlinkInform', 'CrossReferenceSource' );
            newitem.xmlAttributes.add( 'name', link_item.name + stringtype );
            newitem.xmlAttributes.add( 'label', link_item.label + stringtype );
            newitem.xmlAttributes.add( 'crossReferenceFormatIndex', source_item.appliedFormat.index + stringtype );
            //newitem.xmlAttributes.add( 'destinationRelativeURI', destination_item.parent.fullName.getRelativeURI(mydoc.filePath) + stringtype );
            newitem.xmlAttributes.add( 'destinationDocumentName', destination_item.parent.name + stringtype );
            newitem.xmlAttributes.add( 'destinationName', destination_item.name + stringtype );
            //newitem.xmlAttributes.add('destinationId', destination_item.destinationText.associatedXMLElements[0].id + stringtype);
        }
        else if ( source_item.constructor == HyperlinkTextSource )
        {
            if ( destination_item.constructor == HyperlinkURLDestination )
            {
                newitem = source_item.sourceText.associatedXMLElements[0].xmlElements.add( 'xref', source_item.sourceText );
                try
                {
                    newitem.xmlAttributes.add( 'hyperlinkInform', 'HyperlinkTextSource' );
                    newitem.xmlAttributes.add( 'name', link_item.name + stringtype );
                    newitem.xmlAttributes.add( 'label', link_item.label + stringtype );
                    newitem.xmlAttributes.add( 'destinationName', destination_item.name + stringtype );
                    newitem.xmlAttributes.add( 'destinationURL', destination_item.destinationURL + stringtype );
                }
                catch ( err ) { }
            }
        }
    }

    var anchor = null;
    var anchor_count = mydoc.hyperlinkTextDestinations.count();
    for ( var i = 0; i < anchor_count ; i++ )
    {
        anchor = mydoc.hyperlinkTextDestinations[i];
        newitem = null;
        newitem = anchor.destinationText.associatedXMLElements[0].xmlElements.add('xref', anchor.destinationText);
        try { newitem.xmlAttributes.add( 'hyperlinkInform', 'TextAnchor' ); }
        catch ( err ) { }
        try
        {
            newitem.xmlAttributes.add( 'name', anchor.name )
            newitem.xmlAttributes.add( 'label', anchor.label )
        }
        catch ( err )
        { alert('이름 에러\r\nerr', '오류'); }
    }
}

function Set_DestinationID( mydoc )
{
    var xmllist = mydoc.xmlElements[0].evaluateXPathExpression( '//*[@hyperlinkInform]' );
    if ( xmllist.length < 1 )
        return;
    var xml_count = xmllist.length;
    var cross_item = null;
    var hyper_item = null;
    var newitem = null;
    var dest_item = null;
    var dest_doc = null;
    var linktype = null;
    var url = null;
    var dest_file = null;
    for ( var i = 0 ; i < xml_count ; i++ )
    {
        item = xmllist[i];
        linktype = item.xmlAttributes.itemByName( 'hyperlinkInform' ).value;
        if ( linktype == 'CrossReferenceSource' )
        {
            dest_doc = app.open(new File(mydoc.filePath + '/' + item.xmlAttributes.itemByName( 'destinationDocumentName' ).value), false);
            dest_item = dest_doc.hyperlinkTextDestinations.itemByName( item.xmlAttributes.itemByName( 'destinationName' ).value );
            if (!dest_item.isValid)
            {
                alert('상호 참조 오류', '상호 참조 오류');
                exit();
            }
            try
            {
                var para = Get_para_from_char(dest_item.destinationText.associatedXMLElements[0]);
                item.xmlAttributes.add('destinationId', 'd' + para.id);
            }
            catch (err)
            {
                item.xmlAttributes.add('destinationId', '대상 문서에 xml 구조가 없습니다');
            }
        }
    }
}

function Get_para_from_char( c ) {
    if (c.markupTag.name == 'Story' || c.markupTag.name == 'Root' ) {
        alert('구조가 잘못되었습니다');
        exit();
    }
    else if (c.parent.markupTag.name == 'Story' )
        return c;
    else
        return Get_para_from_char(c.parent);
}

//XML구조생성
function NomalStructure( myDoc )
{
    myDoc.xmlPreferences.defaultCellTagName = "Cell";
    myDoc.xmlPreferences.defaultImageTagName = "C_Image";
    myDoc.xmlPreferences.defaultStoryTagName = "Story";
    myDoc.xmlPreferences.defaultTableTagName = "Table";
    var check = 0;
    var myGraphics = myDoc.allGraphics;
    var myGraphicsCount = myGraphics.length;
    var myMaster = myDoc.masterSpreads;
    var myMasterNumber = myMaster.count() - 1;
    var masterImage = 0;
    for ( ; myMasterNumber > -1 ; myMasterNumber-- )
    {
        imageNumber = myMaster[myMasterNumber].allGraphics;
        masterImage += imageNumber.length;
    }
    myGraphicsCount -= masterImage;

    // 태그를 만들 경우, 추가할 스타일을 결정한다

    Temp_XML = new Array();
    var copyXML = {};
    for ( var i = 0; i < myDoc.xmlElements[0].xmlAttributes.count() ; i++ )
    {
        copyXML = new Object();
        copyXML.name = myDoc.xmlElements[0].xmlAttributes[i].name.toString();
        copyXML.value = myDoc.xmlElements[0].xmlAttributes[i].value.toString();
        if ( copyXML.name.search( 'Cell_' ) > -1 )
            Temp_XML = Temp_XML.concat( copyXML );
    }

    DeleteTag( myDoc );         // 스타일 추가전에 구조를 지워준다
    check = CheckTags( myDoc, 0, 1 );
    if ( check < 0 )
        return -1;
    ChangeTagColor( myDoc );
    if (trimCharacterStyleContents == true)
        TrimCharacterStyle(myDoc);
    try {
        myDoc.mapStylesToXMLTags();  // 구조 등록
        myDoc.mapStylesToXMLTags();  // 이미지를 위해 한번 더 구조 등록
    } catch ( err ) {
        errorCount++;
        alert( err, myDoc.name );
        check_CharacterError += err;
        ExportErrorLog( myDoc );
        return 0;
    }
    for ( i = 0 ; i < myGraphicsCount ; i++ ) {
        try { myGraphics[i].autoTag(); }
        catch ( err ) { continue; }
    }
    //specials_to_pi(myDoc);
    if (doNotCheckErrors != true)
        MakeCharErrorLog( myDoc, myDoc.xmlElements[0].xmlElements, myDoc.xmlElements[0].xmlElements.count(), false );
    MappingToStyle( myDoc );
    if ( errorCount ) {
        ExportErrorLog( myDoc );
        //pi_to_breaks_and_specials(myDoc);
        return 0;
    }

    DeleteTag( myDoc );         // 스타일 추가전에 구조를 지워준다
    check = CheckTags( myDoc, 1, 1 );
    if ( check < 0 ) {
        //pi_to_breaks_and_specials(myDoc);
        return -1;
    }
    ChangeTagColor( myDoc );
    myDoc.mapStylesToXMLTags();  // 구조 등록
    myDoc.mapStylesToXMLTags();  // 이미지를 위해 한번 더 구조 등록
    for ( i = 0 ; i < myGraphicsCount ; i++ )
    {
        try { myGraphics[i].autoTag(); }
        catch ( err )
        { continue; }
    }

    //specials_to_pi(myDoc);
    Get_HyperlinkInform( myDoc );
    if (doNotCheckErrors != true)
        MakeCharErrorLog( myDoc, myDoc.xmlElements[0].xmlElements, myDoc.xmlElements[0].xmlElements.count(), true );
    MappingToStyle( myDoc );
    if ( errorCount )
    {
        ExportErrorLog( myDoc );
        //pi_to_breaks_and_specials(myDoc);
        return 0;
    }
    AddImageSize( myDoc );
    SetPageBreakAttr(myDoc);
    //pi_to_breaks_and_specials(myDoc);

    return 1;
}

//단계적으로 XML구조생성
function TreeStructure( myDoc )
{
    myDoc.save();
    var myPath = myDoc.filePath;
    var myFile = myDoc.fullName;
    var myFolder = new Folder( myPath + '/backup2' );
    myFolder.create();
    var myNewFile = new File( myFolder + '/' + myDoc.name.split( '.' )[0] + '_' + CallDate() + '.indd' );
    myFile.copy( myNewFile );

    myDoc.xmlPreferences.defaultCellTagName = "Cell";
    myDoc.xmlPreferences.defaultImageTagName = "C_Image";
    myDoc.xmlPreferences.defaultStoryTagName = "Story";
    myDoc.xmlPreferences.defaultTableTagName = "Table";
    var check = 0;
    var myGraphics = myDoc.allGraphics;
    var myGraphicsCount = myGraphics.length;
    var myMaster = myDoc.masterSpreads;
    var myMasterNumber = myMaster.count() - 1;
    var masterImage = 0;
    for ( ; myMasterNumber > -1 ; myMasterNumber-- )
    {
        imageNumber = myMaster[myMasterNumber].allGraphics;
        masterImage += imageNumber.length;
    }
    myGraphicsCount -= masterImage;

    Temp_XML = new Array();
    var copyXML = {};
    for ( var i = 0; i < myDoc.xmlElements[0].xmlAttributes.count() ; i++ )
    {
        copyXML = new Object();
        copyXML.name = myDoc.xmlElements[0].xmlAttributes[i].name.toString();
        copyXML.value = myDoc.xmlElements[0].xmlAttributes[i].value.toString();
        if ( copyXML.name.search( 'Cell_' ) > -1 )
            Temp_XML = Temp_XML.concat( copyXML );
    }

    DeleteTag( myDoc );         // 스타일 추가전에 구조를 지워준다
    check = CheckTags( myDoc, 0, 1 );
    if ( check < 0 )
        return -1;
    ChangeTagColor( myDoc );
    if (trimCharacterStyleContents == true)
        TrimCharacterStyle(myDoc);
    try {
        myDoc.mapStylesToXMLTags();  // 구조 등록
        myDoc.mapStylesToXMLTags();  // 이미지를 위해 한번 더 구조 등록
    } catch ( err ) {
        errorCount++;
        alert( err, myDoc.name );
        check_CharacterError += err;
        ExportErrorLog( myDoc );
        return 0;
    }
    for ( i = 0 ; i < myGraphicsCount ; i++ ) {
        try { myGraphics[i].autoTag(); }
        catch ( err ) { continue; }
    }

    //specials_to_pi(myDoc);
    if (doNotCheckErrors != true)
        MakeCharErrorLog( myDoc, myDoc.xmlElements[0].xmlElements, myDoc.xmlElements[0].xmlElements.count(), false );
    if ( errorCount ) {
        ExportErrorLog( myDoc );
        //pi_to_breaks_and_specials(myDoc);
        return 0;
    }

    DeleteTag( myDoc );         // 스타일 추가전에 구조를 지워준다
    check = CheckTags( myDoc, 1, 1 );
    if ( check < 0 ) {
        //pi_to_breaks_and_specials(myDoc);
        return -1;
    }
    ChangeTagColor( myDoc );
    myDoc.mapStylesToXMLTags();  // 구조 등록
    myDoc.mapStylesToXMLTags();  // 이미지를 위해 한번 더 구조 등록
    for ( i = 0 ; i < myGraphicsCount ; i++ ) {
        try { myGraphics[i].autoTag(); }
        catch ( err ) { continue; }
    }

    //specials_to_pi(myDoc);
    RemakeStructure( myDoc );
    Get_HyperlinkInform(myDoc);

    if (doNotCheckErrors != true)
        MakeCharErrorLog( myDoc, myDoc.xmlElements[0].xmlElements, myDoc.xmlElements[0].xmlElements.count(), true );
    MappingToStyle(myDoc);
    if ( errorCount ) {
        ExportErrorLog( myDoc );
        //pi_to_breaks_and_specials(myDoc);
        return 0;
    }
    AddImageSize( myDoc );
    SetPageBreakAttr(myDoc);
    //pi_to_breaks_and_specials(myDoc);
    
    return 1;
}

//페이지 넘김 문자 특성 설정
function SetPageBreakAttr(mydoc) {
    try {
        grep_setting("~M");
        var pbs = mydoc.findGrep();
        for (var i = 0 ; i < pbs.length ; i++)
            pbs[i].associatedXMLElements[0].xmlAttributes.add('paragraphBreakType', 'NextColumn');
        return 1;
    } catch (ex) {
        alert('' + ex, 'grep_setting 함수 오류');
        return -1;
    }
}

//특정 태그의 색깔을 지정해주는 함수
function ChangeTagColor( myDoc ) {
    var myTag = myDoc.xmlTags;
    var TagList = myDoc.xmlTags.count() - 1;
    var TagColor = 0;
    var cTagColor = 0;
    var TagName = 0;
    for ( ; TagList > -1 ; TagList-- ) {
        TagName = myTag[TagList].name;
        TagColor = myTag[TagList].tagColor;
        try {
            if ( !( TagColor != 1767336052 ) || ( !( TagColor[0] != 255 ) && !( TagColor[1] != 255 ) && !( TagColor[2] != 255 ) ) )
                myDoc.xmlTags[TagList].tagColor = [RandomNumber( 22, 222 ), RandomNumber( 22, 222 ), RandomNumber( 22, 222 )];
        } catch ( err ) {
            if ( !( TagColor != 'WITHE' ) )
                myDoc.xmlTags[TagList].tagColor = [RandomNumber( 22, 222 ), RandomNumber( 22, 222 ), RandomNumber( 22, 222 )];
        }
        if ( !( TagName != 'MMI' ) || !( TagName != 'MMI_Cover' ) )
            myDoc.xmlTags[TagList].tagColor = 1766680430;
        else if ( !( TagName != 'C_Image' ) )
            myDoc.xmlTags[TagList].tagColor = 1767007588;
        else if ( !( TagName != 'C_Singlestep' ) || !( TagName != 'C_SingleStep' ) )
            myDoc.xmlTags[TagList].tagColor = 1765960821;
    }
}

// 본문에서 호출
// 추가할 태그가 있는지 비교한 뒤에 추가
function CheckTags( myDoc, SelectParagraph, SelectCharacter )
{
    var check = 0;
    if ( SelectParagraph )
    {
        //  단락 스타일의 숫자만큼 체크. 배열숫자 때문에 1을 빼줌
        var TagList = myDoc.xmlTags.everyItem();  // 태그 리스트를 저장
        var TagCounter = myDoc.xmlTags.count() - 1;
        var myPaStyles = myDoc.paragraphStyles;
        check = FindIncorrectStyle( myDoc, myPaStyles );
        if ( check < 0 )
            return -1;

        for ( var myPaCounter = myPaStyles.count() - 1 ; myPaCounter > 1 ; myPaCounter-- )
        {
            myPaStyle = myPaStyles[myPaCounter];        // 단락 스타일
            check = AddToTags( myDoc, myPaStyle );        // 태그와 단락을 비교, 추가
            if ( check < 0 )
                return -1;
        }
        // 스타일 그룹을 위한 추가 코드
        var myPaGroups = myDoc.paragraphStyleGroups;
        for ( var myPaGroupCounter = myPaGroups.count() - 1 ; myPaGroupCounter > 1 ; myPaGroupCounter-- )
        {
            myPaGroup = myPaGroups[myPaGroupCounter];
            var myPaGroupStyles = myPaGroup.paragraphStyles;    // 문자 스타일 그룹
            for ( var myPaCounter = myPaGroupStyles.count() - 1 ; myPaCounter > 1 ; myPaCounter-- )
            {
                myPaGroupStyle = myPaGroupStyles[myPaCounter];
                check = AddToTags( myDoc, myPaGroupStyle );       // 태그와 문자를 비교, 추가
                if ( check < 0 )
                    return -1;
            }
        }
    }

    if ( SelectCharacter )
    {
        var TagList = myDoc.xmlTags.everyItem();  // 태그 리스트를 저장
        var TagCounter = myDoc.xmlTags.count() - 1;
        var myChStyles = myDoc.characterStyles;      // 문자 스타일
        for ( var myChCounter = myChStyles.count() - 1 ; myChCounter > 0 ; myChCounter-- )       //  배열 때문에 -1
        {
            myChStyle = myChStyles[myChCounter];
            check = AddToTags( myDoc, myChStyle );        // 태그와 문자를 비교
            if ( check < 0 )
                return -1;
        }
        // 스타일 그룹을 위한 추가 코드
        var myChGroups = myDoc.characterStyleGroups;
        for ( var myChGroupCounter = myChGroups.count() - 1 ; myChGroupCounter > 0 ; myChGroupCounter-- )
        {
            myChGroup = myChGroups[myChGroupCounter];
            var myChGroupStyles = myChGroup.characterStyles;    // 문자 스타일 그룹
            for ( var myChCounter = myChGroupStyles.count() - 1 ; myChCounter > 0 ; myChCounter-- )
            {
                myChGroupStyle = myChGroupStyles[myChCounter];
                check = AddToTags( myDoc, myChGroupStyle );       // 태그와 문자를 비교
                if ( check < 0 )
                    return -1;
            }
        }
    }
    return 1;
}

// 스타일을 태그에 추가해주는  함수
function AddToTags( myDoc, myStyle ) {
    try {
        try { var myTagName = myStyle.name; }
        catch ( err ) { alert( myStyle.name + '이 정상적인 스타일이 아닙니다', '스타일오류' ); }
        // Make a Tag with the Style name..
        try { myDoc.xmlTags.add( { name: myTagName } ); }
        catch ( err ) {    // Tag가 있는 경우 충돌을 핸들링
            try { myDoc.xmlTags.itemByName( myTagName ).index; }
            catch ( err ) {
                errorCount++; styleErrorCount++;
                alert( myDoc.name.split( '.' )[0] + ' 의 ' + myStyle.name + '은 태그화 할 수 있는 정상적인 스타일이 아닙니다', '스타일오류' );
                return -1;
            }
        }
        myDoc.xmlExportMaps.add( myTagName, myTagName );
    } catch ( err ) {
        alert( 'AddToTags 함수 에러', '오류' );
        return -1;
    }
    return 1;
}

// 문서의 단락 스타일에 [없음]과 [기본단락]이 사용되었는지 찾는 함수
function FindIncorrectStyle( myDoc, myParagraphStyle )
{
    app.findTextPreferences = NothingEnum.NOTHING;
    app.findTextPreferences.appliedParagraphStyle = myParagraphStyle[1];
    var detect = app.findText();
    var detect_Length = detect.length;
    if ( detect_Length )
    {
        errorCount++; styleErrorCount++;
        alert( myParagraphStyle[1].name + '이 사용되었습니다', '스타일오류' );
        app.findTextPreferences = NothingEnum.NOTHING;
        return -1;
    }
    app.findTextPreferences.appliedParagraphStyle = myParagraphStyle[0];
    detect = app.findText();
    detect_Length = detect.length;
    if ( detect_Length )
    {
        errorCount++; styleErrorCount++;
        alert( myParagraphStyle[0].name + '이 사용되었습니다', '스타일오류' );
        app.findTextPreferences = NothingEnum.NOTHING;
        return -1;
    }
    app.findTextPreferences = NothingEnum.NOTHING;
    return 1;
}

// 본문에서 호출
// 단락 스타일을 태그에 생성해주는  함수
function AddParagraphStyles( myDoc )
{
    var check = 0;
    var Mycounter = myDoc.paragraphStyles.count() - 1;
    var myParagraphStyle = myDoc.paragraphStyles;
    check = FindIncorrectStyle( myDoc, myParagraphStyle );
    if ( check < 0 )
        return -1;

    for ( ; Mycounter > 4  ; )
    {
        check = AddToTags( myDoc, myParagraphStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
        check = AddToTags( myDoc, myParagraphStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
        check = AddToTags( myDoc, myParagraphStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
    }
    for ( ; Mycounter > 1  ; )
    {
        check = AddToTags( myDoc, myParagraphStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
    }
    check = AddParagraphGroup( myDoc );
    if ( check < 0 )
        return -1;
    return 1;
}

// 본문에서 호출
// 문자 스타일을 태그에 생성해주는 함수
function AddCharacterStyles( myDoc )
{
    var check = 0;
    var Mycounter = myDoc.characterStyles.count() - 1;
    var myCharacterStyle = myDoc.characterStyles;
    for ( ; Mycounter > 3  ; )
    {
        check = AddToTags( myDoc, myCharacterStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
        check = AddToTags( myDoc, myCharacterStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
        check = AddToTags( myDoc, myCharacterStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
    }
    for ( ; Mycounter > 0  ; )
    {
        check = AddToTags( myDoc, myCharacterStyle[Mycounter--] );
        if ( check < 0 )
            return -1;
    }
    check = AddCharacterGroup( myDoc );
    if ( check < 0 )
        return -1;
    return 1;
}

// AddParagraphStyles에서 호출
// 그룹안의 단락 스타일을 태그에 생성해주는 함수
function AddParagraphGroup( myDoc )
{
    var check = 0;
    var myPaGroupCounter = myDoc.paragraphStyleGroups.count() - 1;
    for ( ; myPaGroupCounter > -1  ; myPaGroupCounter-- )
    {
        var myPaGroup = myDoc.paragraphStyleGroups[myPaGroupCounter];
        var myPaCounter = myPaGroup.paragraphStyles.count() - 1;
        for ( ; myPaCounter > 4 ; )
        {
            check = AddToTags( myDoc, myPaGroup.paragraphStyles[myPaCounter--] );
            if ( check < 0 )
                return -1;
            check = AddToTags( myDoc, myPaGroup.paragraphStyles[myPaCounter--] );
            if ( check < 0 )
                return -1;
            check = AddToTags( myDoc, myPaGroup.paragraphStyles[myPaCounter--] );
            if ( check < 0 )
                return -1;
        }
        for ( ; myPaCounter > 1 ; )
        {
            check = AddToTags( myDoc, myPaGroup.paragraphStyles[myPaCounter--] );
            if ( check < 0 )
                return -1;
        }
    }
    return 1;
}

// AddCharacterStyles에서 호출
// 그룹안의 단락 스타일을 태그에 생성해주는 함수
function AddCharacterGroup( myDoc )
{
    var check = 0;
    var myChGroupCounter = myDoc.characterStyleGroups.count() - 1;
    for ( ; myChGroupCounter > -1  ; myChGroupCounter-- )
    {
        var myChGroup = myDoc.characterStyleGroups[myChGroupCounter];
        var myChCounter = myChGroup.characterStyles.count() - 1;
        for ( ; myChCounter > 3 ; )
        {
            check = AddToTags( myDoc, myChGroup.characterStyles[myChCounter--] );
            if ( check < 0 )
                return -1;
            check = AddToTags( myDoc, myChGroup.characterStyles[myChCounter--] );
            if ( check < 0 )
                return -1;
            check = AddToTags( myDoc, myChGroup.characterStyles[myChCounter--] );
            if ( check < 0 )
                return -1;
        }
        for ( ; myChCounter > 0 ; )
        {
            check = AddToTags( myDoc, myChGroup.characterStyles[myChCounter--] );
            if ( check < 0 )
                return -1;
        }
    }
    return 1;
}

// 본문에서 호출
// 구조를 지워주는 함수
function StructureRemove( myDoc )
{
    var count = myDoc.xmlElements[0].xmlElements.count();
    var unTagSection = myDoc.xmlElements[0].xmlElements;
    for ( ; count > 0 ; count-- )
        unTagSection[0].untag();
}

function DeleteTag( myDoc ) {

    //pi_to_breaks_and_specials(myDoc);
    StructureRemove( myDoc );
    myDoc.deleteUnusedTags();
    myDoc.xmlElements[0].remove();
}

// 본문에서 호출
// XML을 불러오는 함수
function ImportXML( myDoc, sourceFile, myImageSize )
{
    errorCount = 0;
    var failList = '';
    var check = 0;

    myDoc.xmlImportPreferences.importStyle = 1481469289; //'XMmi'
    myDoc.xmlImportPreferences.createLinkToXML = false;
    myDoc.xmlImportPreferences.allowTransform = false;
    myDoc.xmlImportPreferences.repeatTextElements = false;
    myDoc.xmlImportPreferences.ignoreUnmatchedIncoming = false;
    myDoc.xmlImportPreferences.importTextIntoTables = false;
    myDoc.xmlImportPreferences.ignoreWhitespace = true;
    myDoc.xmlImportPreferences.removeUnmatchedExisting = true;
    myDoc.xmlImportPreferences.importCALSTables = true;
    myDoc.xmlImportPreferences.importToSelected = false;
    app.pdfPlacePreferences.pdfCrop = 1131573313; // art

    try
    {
        try { var myPath = myDoc.filePath; }
        catch ( err )
        {
            resultLog += myDoc.name + ' 가 한 번도 저장된 적이 없습니다\r\n';
            totalError++; return -1;
        }
        if ( myDoc.xmlElements[0].xmlElements.count() < 1)
        {
            resultLog += myDoc.name.split( '.' )[0] + '\r\n' + myDoc.name.split( '.' )[0] + ' 문서가 XML구조를 가지고 있지 않습니다\r\n';
            if (sourceFile != 0)
                IndexErrorLog( myDoc, myDoc.name.split( '.' )[0] + ' 문서가 XML구조를 가지고 있지 않습니다\r\n' );
            totalError++;
            return -1;
        }
        var inddFile = new File( myPath + '/XML/' + myDoc.name.split('.')[0] + '.xml' );
        if (inddFile.exists) {
            try {
                var xmlFile1 = new File(myPath + '/XML/' + myDoc.name.split('.')[0] + '.xml');
                xmlFile1.open('r');
                var xml1 = new XML(xmlFile1.read());
                xmlFile1.close();
                var donot = xml1.xpath("/Root[@doNotErrorChecking]");
                if (donot != null && donot.toString() != null && donot.toString().length > 0) {
                    resultLog += xmlFile1.name + ' 문서가 오류 확인 기능을 제외한 상태에서 작성된 문서입니다. 해당 XML은 불러올 수 없습니다.\r\n';
                    totalError++; return -1;
                } else
                    myDoc.importXML( inddFile );
            } catch ( err ) {
                resultLog += myDoc.name.split( '.' )[0] + ' 문서에 링크 오류가 있거나 XML파일이 손상되었습니다\r\n';
                totalError++; return -1;
            }
        } else {
            resultLog += myDoc.name.split( '.' )[0] + ' 파일 이름과 똑같은 XML 파일이 없습니다\r\n';
            totalError++; return -1;
        }

        Delete_Hyperlink(myDoc);
        ResizeImage( myDoc, myImageSize );
        //specials_to_pi(myDoc);
        MappingToStyle( myDoc );
        //pi_to_breaks_and_specials(myDoc);
        if ( myImageSize == 'UM'  && sourceFile != 0 )
            check = MakeIndex( myDoc, sourceFile );
        chkSwatchName( "ChangedMMI", myDoc );
        chkSwatchName( "ChangedMMI2", myDoc );
        AddChangedStyle( myDoc );
        Changed_MMI( myDoc, "MMI" );
        Changed_MMI( myDoc, "MMI_NoBold" );
        
		findchangeMMI ( myDoc );
		updateCross ( myDoc );
    }
    catch ( err )
    {
        failList += myDoc.name + '\r\n' + check + '\r\n';
        errorCount++;
    }
    if ( errorCount )
    {
        resultLog += '\r\n' + failList;
        totalError++;
        return -1;
    }
    else if ( check )
    {
        indexingLog += '\r\n' + check;
        indexError++;
        return -1;
    }
    else
        return 1;
}

// 하이퍼링크 생성
function Set_HyperlinkInform( mydoc )
{
    var xmllist = mydoc.xmlElements[0].evaluateXPathExpression( '//*[@hyperlinkInform]' );
    if ( xmllist.length < 1 )
        return;
    var xml_count = xmllist.length;
    var cross_item = null;
    var hyper_item = null;
    var newitem = null;
    var dest_item = null;
    var dest_doc = null;
    var linktype = null;
    var url = null;
    var dest_file = null;
    for ( var i = 0 ; i < xml_count ; i++ )
    {
        item = xmllist[i];
        linktype = item.xmlAttributes.itemByName( 'hyperlinkInform' ).value;
        if ( linktype == 'CrossReferenceSource' )
        {
            cross_item = mydoc.crossReferenceSources.add( item.texts[0], mydoc.crossReferenceFormats[item.xmlAttributes.itemByName( 'crossReferenceFormatIndex' ).value * 1] );
            //dest_doc = app.open(new File(mydoc.filePath + '/' + item.xmlAttributes.itemByName( 'destinationRelativeURI' ).value), false);
            dest_file = new File(mydoc.filePath + '/' + item.xmlAttributes.itemByName( 'destinationDocumentName' ).value);
            if (!dest_file.exists)
            {
                alsert(dest_file + ' 문서가 없습니다', '파일 오류');
                exit();
            }
            dest_doc = app.open(new File(mydoc.filePath + '/' + item.xmlAttributes.itemByName( 'destinationDocumentName' ).value), false);
            dest_item = dest_doc.hyperlinkTextDestinations.itemByName( item.xmlAttributes.itemByName( 'destinationName' ).value );
            if (!dest_item.isValid)
            {
                alert('상호 참조 오류', '상호 참조 오류');
                exit();
            }
            newitem = mydoc.hyperlinks.add( cross_item, dest_item );
            newitem.highlight =HyperlinkAppearanceHighlight.INVERT;
        }
        else if ( linktype == 'HyperlinkTextSource' )
        {
            hyper_item = mydoc.hyperlinkTextSources.add( item.texts[0] );
            url = mydoc.hyperlinkURLDestinations.itemByName( item.xmlAttributes.itemByName( 'name' ).value );
            if ( !url.isValid || url.destinationURL != item.xmlAttributes.itemByName( 'destinationURL' ).value )
                url = mydoc.hyperlinkURLDestinations.add( item.xmlAttributes.itemByName( 'destinationURL' ).value, { hidden:true });
            newitem = mydoc.hyperlinks.add( hyper_item, url , { visible:false, hidden:false });
            try
            {
                newitem.name = item.xmlAttributes.itemByName( 'name' ).value;
                newitem.label = item.xmlAttributes.itemByName( 'label' ).value;
            }
            catch ( err ) { }
        }
    }
    return 1;
}

function GetAddress( myXmlElementItem )
{
    var index = [];
    var j = 0;
    do
    {
        index[j++] = myXmlElementItem.index;
        myXmlElementItem = myXmlElementItem.parent;
    } while ( myXmlElementItem.markupTag.name != 'Root' )
    index.reverse();
    return index;
}

function GetXMLitem( mydoc, index )
{
    var item = mydoc.xmlElements[0];
    for ( var i = 0 ; i < index.length ; i++ )
    {
        item = item.xmlElements[index[i] * 1];
    }
    return item;
}

// 이미지의 크기를 일괄설정 해주는 함수
function ResizeImage( myDoc, imageScale )
{
    if ( imageScale < 0 )
        return 0;
    try
    {
        var allImages = myDoc.allGraphics;
        var myMaster = myDoc.masterSpreads;
        var myMasterNumber = myMaster.count() - 1;
        var masterImage = 1;
        for ( ; myMasterNumber > -1 ; myMasterNumber-- )
        {
            imageNumber = myMaster[myMasterNumber].allGraphics;
            masterImage += imageNumber.length;
        }
        var imageCount = allImages.length - masterImage;
        if ( ( imageScale == 'UM' ) || ( imageScale == 'QSG' ) || ( imageScale == 'Leaflet' ) )
            for ( ; imageCount > -1 ; imageCount-- )
            {
                try
                {
                    try
                    {
                        allImages[imageCount].absoluteHorizontalScale = +allImages[imageCount].associatedXMLElement.xmlAttributes.itemByName('xScale').value;
                        allImages[imageCount].absoluteVerticalScale = +allImages[imageCount].associatedXMLElement.xmlAttributes.itemByName('yScale').value;
                    }
                    catch (err) {}
                    allImages[imageCount].fit( FitOptions.CENTER_CONTENT );
                    allImages[imageCount].fit( FitOptions.FRAME_TO_CONTENT );
                } catch ( err ) { }
            }
        else
            for ( ; imageCount > -1 ; imageCount-- )
            {
                try
                {
                    allImages[imageCount].absoluteVerticalScale = imageScale;
                    allImages[imageCount].absoluteHorizontalScale = imageScale;
                    allImages[imageCount].fit( FitOptions.CENTER_CONTENT );
                    allImages[imageCount].fit( FitOptions.FRAME_TO_CONTENT );
                } catch ( err ) { }
            }
    } catch ( err ) { alert( 'ResizeImage 함수 에러', '오류' ); }
}

// 문서의 XML구조에 이미지들의 크기를 특성값으로 넣어주는 함수
function AddImageSize( myDoc )
{
    try
    {
        var allImages = myDoc.allGraphics;
        var myMaster = myDoc.masterSpreads;
        var myMasterNumber = myMaster.count() - 1;
        var masterImage = 1;
        for ( ; myMasterNumber > -1 ; myMasterNumber-- )
        {
            imageNumber = myMaster[myMasterNumber].allGraphics;
            masterImage += imageNumber.length;
        }
        var imageCount = allImages.length - masterImage;
        for ( ; imageCount > -1 ; imageCount-- )
        {
            try
            {
                allImages[imageCount].associatedXMLElement.xmlAttributes.add( 'xScale', allImages[imageCount].absoluteHorizontalScale.toString() );
                allImages[imageCount].associatedXMLElement.xmlAttributes.add( 'yScale', allImages[imageCount].absoluteVerticalScale.toString() );
            } catch ( err ) { }
        }
    } catch ( err ) { alert( 'addImageSize 함수 에러', '오류' ); }
}

// XML 구조를 호출 할 경우 문서의 종류를 정해주는 함수
function SelectManualType()
{
    var res =
        "dialog { alignChildren: 'fill', \
            Type: Panel { orientation: 'column', alignChildren:'left', \
                text: '메뉴얼 스타일',\
                SelectType: Group { orientation: 'row', \
                    UM: Button { text: 'BOOK-UM / ONLINE', minimumSize: [80, 25] },\
                    QSG: Button { text: 'BOOK-QSG', minimumSize: [80, 25] },\
                    Leaflet: Button { text: 'Leaflet-UM', minimumSize: [80, 25] },\
                }\
            }, \
            Notice: Panel { orientation: 'column', alignChildren:'left', \
                text: '설명', \
                s: StaticText { text: 'BOOK-UM / ONLINE: Image Size 170%'},\
                s: StaticText { text: 'BOOK-QSG: Image Size 100%'},\
                s: StaticText { text: 'Leaflet-UM: Image Size 100%'},\
            },\
            buttons: Group { orientation: 'row', alignment: 'left', \
                backBtn: Button { text:'뒤로', alignment:'left' }, \
            } \
        }";
    var type;
    var select = new Window( res, '문서의 종류를 선택해주세요' );
    select.center();
    select.Type.SelectType.UM.value = true;
    select.Type.SelectType.UM.onClick = function () { type = 'UM'; select.close( 1 ); }
    select.Type.SelectType.QSG.onClick = function () { type = 'QSG'; select.close( 1 ); }
    select.Type.SelectType.Leaflet.onClick = function () { type = 'Leaflet'; select.close( 1 ); }
    select.buttons.backBtn.onClick = function () { select.close( -1 ); }
    select.helpTip = '문서의 종류에 따른 이미지의 크기를 선택합니다';
    select.Type.SelectType.UM.helpTip = 'UM용 이미지 크기로 재설정 합니다';
    select.Type.SelectType.QSG.helpTip = 'QSG용 이미지 크기로 재설정 합니다';
    select.Type.SelectType.Leaflet.helpTip = 'Leaflet용 이미지 크기로 재설정 합니다';
    var result = select.show() - 1;
    if ( result == 0 )
        return type;
    else if ( result < 0 )
        return 0;
    else
        return -1;
}

// 색인을 만들어주는 함수
function MakeIndex( myDoc, sourceFile )
{
    var indexErrors = 0;
    var sourceTagCount = 0; var targetTagCount = 0;
    try
    {
        ResetIndex( myDoc );
        var targetDoc = myDoc;
        CountXML( myDoc, myDoc.xmlElements[0].xmlElements, myDoc.xmlElements[0].xmlElements.count() );
        var targetCount = totalXMLCount;
        totalXMLCount = 0;
        var myDocName = myDoc.name.split( '%20' ).join( ' ' ).split( '.' )[0];
        var errorLog = '';
        if ( myDocName != sourceFile.name.split( '%20' ).join( ' ' ).split( '.' )[0] )
            return IndexErrorLog( targetDoc, '파일을 잘못 불러왔습니다\r\n' );
        var sourceDoc = null;
        try { sourceDoc = app.open( sourceFile ); }
        catch ( err )
        { return IndexErrorLog( targetDoc, '색인을 가져올 인디자인 파일이 없습니다\r\n' ); }
        if ( !sourceDoc.xmlElements[0].xmlElements.count() )
        {
            sourceDoc.close( SaveOptions.NO );
            return IndexErrorLog( targetDoc, '영문 인디자인 파일이 XML구조를 가지고 있지 않습니다\r\n구조화 후 다시 시도해주세요\r\n' );
        }
        CountXML( sourceDoc, sourceDoc.xmlElements[0].xmlElements, sourceDoc.xmlElements[0].xmlElements.count() );
        var sourceCount = totalXMLCount;
        totalXMLCount = 0;
        if ( targetCount != sourceCount ) { sourceDoc.close( SaveOptions.NO );
            return IndexErrorLog( targetDoc, 'XML 구조가 서로 다릅니다\r\n' ); }
        var sourceIndex = sourceDoc.indexes[0];
        var targetIndex = app.documents[1].indexes[0];
        var mySelection; var sourceXML = [];

        var xmlSource = [];
        var sourceTopicLeng = null;
        try { sourceTopicLeng = sourceIndex.allTopics.length; }
        catch ( err )
        {
            ResetIndex( targetDoc );
            sourceDoc.close( SaveOptions.NO );
            return IndexErrorLog( targetDoc, '영문 인디자인 파일에 색인 정보가 없습니다' );
        }
        var sourcePage = 0;
        var tempArray = [];
        for ( var i = 0 ; i < sourceTopicLeng ; i++ )
        {
            try { sourcePage = sourceIndex.allTopics[i].pageReferences[0].sourceText; }
            catch ( err ) { continue; }
            sourceXML = sourcePage.associatedXMLElements[0];
            try
            {
                tempArray = [i, sourceXML];
                sourceXML.markupTag.name;
                xmlSource.push( tempArray );
            }
            catch ( err )
            {
                sourceDoc.close( SaveOptions.NO );
                return IndexErrorLog( targetDoc, '영문 인디자인 파일의 XML구조에 문제가 있습니다\r\n' );
            }
        }
        var xmlSourceLeng = xmlSource.length;
        for ( var i = 0 ; i < xmlSourceLeng ; i++ )
            xmlSource[i][1] = ExportXMLposition( xmlSource[i][1] );
        sourceDoc.close( SaveOptions.NO );
        targetDoc = myDoc;
        targetIndex = myDoc.indexes[0];
        targetIndex.importTopics( sourceFile );
        var targetXML;
        var myRoot = targetDoc.xmlElements[0];
        for ( var i = 0 ; i < xmlSourceLeng ; i++ )
            xmlSource[i][1] = ChangePositionToXML( myRoot, xmlSource[i][1] );
        for ( var i = 0 ; i < xmlSourceLeng ; i++ )
        {
            try
            {
                targetXML = xmlSource[i][1].insertionPoints[0];
                targetPoint = targetXML.index;
                targetIndex.allTopics[xmlSource[i][0]].pageReferences.add( targetXML );
                var check = 0;
                do
                {
                    var realpoint = targetIndex.allTopics[xmlSource[i][0]].pageReferences[0].sourceText.index;
                    if ( Math.abs( realpoint - targetPoint ) < 1 )
                    { check = 0; break; }
                    else if ( realpoint > targetPoint )
                    {
                        targetDoc.undo();
                        try { myDoc.indexes[0].isValid; } catch ( err ) { targetDoc.undo(); }
                        targetXML = targetXML.parent.insertionPoints[targetXML.index - 1];
                        targetIndex.allTopics[xmlSource[i][0]].pageReferences.add( targetXML );
                    }
                    else if ( realpoint < targetPoint )
                    {
                        targetDoc.undo();
                        try { myDoc.indexes[0].isValid; } catch ( err ) { targetDoc.undo(); }
                        targetXML = targetXML.parent.insertionPoints[targetXML.index + 1];
                        targetIndex.allTopics[xmlSource[i][0]].pageReferences.add( targetXML );
                    }
                } while ( check++ < 125 )
            }
            catch ( err )
            {
                var limit = 20;
                do
                {
                    limit--;
                    var targetTopic = targetIndex.allTopics[xmlSource[i][0]];
                    var errorTopicName = targetTopic.name;
                    check = 1;
                    try
                    {
                        targetTopic = targetTopic.parent;
                        targetTopic.indexSections;
                        errorTopicName = '>>' + errorTopicName;
                        check = 0;
                    } catch ( err ) { errorTopicName = targetTopic.name + '>>' + errorTopicName; }
                } while ( check | limit )
                if ( !limit )
                    return IndexErrorLog( targetDoc, 'XML 구조가 서로 맞지 않습니다' );
                errorLog += errorTopicName + '\r\n';
                indexErrors++;
                check = 0;
                continue;
            }
            if ( check )
            {
                targetDoc.undo();
                try { myDoc.indexes[0].isValid; } catch ( err ) { targetDoc.undo(); }
                indexErrors++;
                var targetTopic = targetIndex.allTopics[xmlSource[i][0]];
                var errorTopicName = targetTopic.name;
                check = 1;
                do
                {
                    try
                    {
                        targetTopic = targetTopic.parent;
                        targetTopic.indexSections;
                        check = 0;
                    } catch ( err ) { errorTopicName = targetTopic.name + '>>' + errorTopicName; }
                } while ( check )
                errorLog += errorTopicName + '\r\n';
            }
        }
        if ( indexErrors )
            return IndexErrorLog( targetDoc, errorLog );
    }
    catch ( err ){ return IndexErrorLog( targetDoc, 'MakeIndex 함수 오류' ); }
    return 0;
}

//색인작업을 하기전, 영문 INDD파일과 대상 INDD파일의 XML구조가 일치하는지 알기 위해 숫자를 산출
function CountXML( myDoc, myXmlElements, count )
{
    var ElementsList = count;
    for ( var i = 0 ; i < ElementsList ; i++ )
    {
        totalXMLCount++;
        myXmlElementItem = myXmlElements[i];
        myXmlElementItemElements = myXmlElements[i].xmlElements;
        var recount = myXmlElementItemElements.count();
        if ( recount )
            CountXML( myDoc, myXmlElementItemElements, recount );
    }
}

//색인작업 수행 이전에 대상 INDD파일의 색인을 초기화
function ResetIndex( myDoc )
{
    try
    {
        var myTopics = myDoc.indexes[0].allTopics;
        var myTopicsCount = myTopics.length - 1;
        for ( ; myTopicsCount > -1 ; myTopicsCount-- )
        { try { myTopics[myTopicsCount].remove(); } catch ( err ) { } }
    } catch ( err ) { myDoc.indexes.add(); }
}
// 색인작업 중 발생한 오류 로그관련 함수
function IndexErrorLog( myDoc, errorLog )
{
    try
    {
        var myName = myDoc.name.toString().split( "." );
        var myFolder = new Folder( myDoc.filePath + '/Index' );
        myFolder.create();
    }
    catch ( err )
    {
        alert( myName[0] + ' 문서가 한 번도 저장되지 않았습니다', '오류' );
        totalError++;
        return -1;
    }
    try
    {
        var indexLogFile = new File( myDoc.filePath + '/Index/' + myName[0] + '_IndexLog' + '.txt' );
        indexLogFile.open( "w" );
        indexLogFile.write( errorLog );
        indexLogFile.close();
    }
    catch ( err )
    {
        alert( myName[0] + ' Index Log 작성 실패', '오류' );
        return -1;
    }
    return myDoc.name;
}

//영문 INDD파일 색인의 주소를 positon값으로 뽑아내서 배열값에 넘겨줌
function ExportXMLposition( sourceXML )
{
    var arrayIndex = []; var check = 1;
    do
    {
        if ( sourceXML.markupTag.name != 'Root' )
        {
            arrayIndex.push( sourceXML.index );
            sourceXML = sourceXML.parent;
        }
        else check = 0;
    } while ( check )
    arrayIndex.reverse();
    return arrayIndex;
}
// position값을 XML주소값으로 바꿔주는 함수
function ChangePositionToXML( myRoot, sourcePoint )
{
    var length = sourcePoint.length;
    try
    {
        for ( var i = 0 ; i < length ; i++ )
            myRoot = myRoot.xmlElements[sourcePoint[i]];
        return myRoot;
    } catch ( err ) { return -1; }
}
//XMl파일 생성
function ExportXML( myDoc )
{
    myDoc.xmlExportPreferences.excludeDtd = false;
    myDoc.xmlExportPreferences.viewAfterExport = false;
    myDoc.xmlExportPreferences.exportFromSelected = false;
    myDoc.xmlExportPreferences.exportUntaggedTablesFormat = 1484022643; //CALS
    myDoc.xmlExportPreferences.characterReferences = false;
    myDoc.xmlExportPreferences.allowTransform = false;
    myDoc.xmlExportPreferences.fileEncoding = 1937134904; //UTF-8
    myDoc.xmlExportPreferences.ruby = false;

    var myName = null;
    var myFolder = null;
    try
    {
        myName = myDoc.name.toString().split( "." );
        myFolder = new Folder( myDoc.filePath + '/XML' );
        myFolder.create();
    }
    catch ( err )
    {
        alert( myName[0] + ' 문서가 한 번도 저장되지 않았습니다', '오류' );
        totalError++;
        return -1;
    }
    var FileOutput = new File(myDoc.filePath + '/XML/' + myName[0] + '.xml');
    if (doNotCheckErrors)
        FileOutput = new File(myDoc.filePath + '/XML/' + myName[0] + '_DonotErrorCheck.xml');
    //specials_to_pi(myDoc);
    try
    {
        var indent = new File((new File($.fileName)).parent +  '/indent_XML.xsl');
        if (!indent.exists)
        {
            var vbscript = '''version()
                        Function version
                            Dim myURL
                            myURL = "http://download.astkorea.net/InDesignScript/indent_XML.xsl"
                            Dim WinHttpReq
                            Set WinHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
                            WinHttpReq.setOption 2, 13056
                            WinHttpReq.open "GET", myURL, False
                            WinHttpReq.send ""
                            
                            set objStream = CreateObject("ADODB.Stream")
                            objStream.Type = 1 'adTypeBinary
                            objStream.Open
                            objStream.Write WinHttpReq.responseBody
                            objStream.SaveToFile "''' + sFolder_fs + '''\\indent_XML.xsl", 2
                            objStream.Close
                            set objStream = Nothing

                        End Function''';
            app.doScript (vbscript, ScriptLanguage.VISUAL_BASIC); 
        }
        myDoc.xmlExportPreferences.allowTransform = true;
        myDoc.xmlExportPreferences.transformFilename = indent;
        myDoc.exportFile( ExportFormat.xml, FileOutput );
    }
    catch ( err )
    {
        myDoc.xmlExportPreferences.allowTransform = false;
        myDoc.exportFile( ExportFormat.xml, FileOutput );
    }
    if ( myXSLT )
    {
        myDoc.xmlExportPreferences.allowTransform = true;
        myDoc.xmlExportPreferences.transformFilename = myXSLT;
        FileOutput = new File( myDoc.filePath + '/XML/' + myName[0] + '_MMI.xml' );
        myDoc.exportFile( ExportFormat.xml, FileOutput );
        myDoc.xmlExportPreferences.allowTransform = false;
    }
    //pi_to_breaks_and_specials(myDoc);

    return 1;
}

// 본문에서 호출
// 에러로그를 생성하는 함수
function ExportErrorLog( myDoc )
{
    var myName = null;
    var myFolder = null;
    try
    {
        myName = myDoc.name.toString().split( "." );
        myFolder = new Folder( myDoc.filePath + '/Log' );
        myFolder.create();
    }
    catch ( err )
    {
        alert( myName[0] + ' 문서가 한 번도 저장되지 않았습니다', '오류' );
        totalError++;
        return -1;
    }
    try
    {
        var FileOutput = new File( myDoc.filePath + '/Log/' + myName[0] + '_Log' + '.txt' );
        FileOutput.open( "w" );
        FileOutput.write( check_CharacterError );
        FileOutput.close();
        if ( FileOutput.exists )
        {
            FileOutput.open( "e" );
            check_CharacterError = FileOutput.read();
            var reCheck_CharacterError = check_CharacterError.split( '?' ).join( '' );
            FileOutput.close();
            FileOutput.open( "w" );
            FileOutput.write( reCheck_CharacterError );
            FileOutput.close();
        }
    } catch ( err ) { alert( 'ExportErrorLog 함수 에러', '오류' ); }
    check_CharacterError = 'CharacterStyle Error Log' + '\r\n\r\n';
    write_ErrorAddress = '';
}

// 본문에서 호출
// 구조를 단계화 해준다
function RemakeStructure( myDoc )
{
    this.name = "RemakeStructure";
    this.toString = function () { return this.name; }

    var rootSection = myDoc.xmlElements[0].xmlElements;
    var myRoot = myDoc.xmlElements[0];
    MakeParent( myDoc );
    var myStoryCount = myRoot.xmlElements.count() - 1;
    HeadingCounter = CheckHeading( myDoc );
    for ( ; myStoryCount > -1 ; myStoryCount-- )
    {
        myStory = myRoot.xmlElements[myStoryCount].xmlElements;
        MakeTopic( myDoc, myStory, 1, HeadingCounter );
    }
    MakeHead( myDoc );
    MakeTopicInHead( myDoc );
    MappingToStyle( myDoc );
}

// 구조를 단계화 할때 Heading의 값을 정해주는 함수
function CheckHeading( myDoc )
{
    var check = 0;
    var checkName = null;
    var TagList = myDoc.xmlTags;
    var TagCounter = myDoc.xmlTags.count() - 1;
    var i = 1;
    try
    {
        for ( ; i < TagCounter ; i++ )
        {
            check = 0;
            checkName = 'Heading' + i;
            try
            {
                TagList.itemByName( checkName ).name;
                check |= 1;
            } catch ( err ) { }
            checkName = 'Heading_' + i;
            try
            {
                TagList.itemByName( checkName ).name;
                check |= 1;
            } catch ( err ) { }
            checkName = 'SafetyHeading' + i;
            try
            {
                TagList.itemByName( checkName ).name;
                check |= 1;
            } catch ( err ) { }
            if ( !check )
                break;
        }
        return --i;
    } catch ( err ) { alert( 'CheckHeading 함수 에러', '오류' ); }
}

// RemakeStructure 에서 호출
// Parent 요소를 만들 조건을 체크한다.
function MakeParent( myDoc )
{
    checkTagName = 0;
    TagNameLength = 0;
    checkTagName2 = 0;
    var rootSection = myDoc.xmlElements[0].xmlElements;
    var rootSectionCounter = rootSection.count() - 1;
    try
    {
        for ( ; rootSectionCounter > -1 ; rootSectionCounter-- )
        {
            var storySection = rootSection[rootSectionCounter].xmlElements;
            var storySectionCounter = storySection.count() - 1;
            if ( storySectionCounter )
                for ( ; storySectionCounter > 0 ; )
                {
                    TagName = storySection[storySectionCounter--].markupTag.name;
                    checkTagName = TagName.split( "_" );
                    TagNameLength = checkTagName[0].length;
                    checkTagName2Full = storySection[storySectionCounter].markupTag.name;
                    checkTagName2 = checkTagName2Full.substring( 0, TagNameLength );

                    if ( ( !( checkTagName[0] != 'Empty' ) || !( checkTagName[0] != 'Contents' ) || !( checkTagName[0] != 'Topic' ) ) || ( ( checkTagName[0].search( 'Heading' ) > -1 ) || ( TagName.search( 'Title' ) > -1 ) || ( checkTagName[0].search( 'PA_' ) > -1 ) ) ); // 해당일 경우 패스
                    else
                    {
                        if ( !( checkTagName[0] != checkTagName2 ) || !( checkTagName2Full != 'Empty' ) )
                        {
                            startPoint = storySectionCounter + 1;
                            while ( !( checkTagName2Full != 'Empty' ) && storySectionCounter )
                                checkTagName2Full = storySection[--storySectionCounter].markupTag.name;
                            checkTagName2 = checkTagName2Full.substring( 0, TagNameLength );
                            if ( !( checkTagName[0] != checkTagName2 ) )
                            {
                                styleName = myDoc.paragraphStyles.itemByName( TagName );
                                check = 1;
                                AddElement( storySection, 'PA_' + checkTagName[0], startPoint );
                                arrayPoint = startPoint + 1;
                                storySectionCounter = MakePA( storySection, storySectionCounter );
                            }
                        }
                    }
                }
        }
    } catch ( err ) { alert( 'MakeParent 함수 에러', '오류' ); }
}

// MakeParent에서 호출
// PA_를 만들어서 ssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssssss같은 계열의 연속된 요소를 묶어준다.
function MakePA( storySection, storySectionCounter )
{
    var checkEmpty = 0;
    var checkDiff = 0;
    try
    {
        do
        {
            checkTagName2Full = storySection[--storySectionCounter].markupTag.name;
            checkTagName2 = checkTagName2Full.substring( 0, TagNameLength );
            if ( !storySectionCounter )
            {
                check = 0;
                if ( checkTagName[0] != checkTagName2 )
                {
                    endPoint = storySectionCounter + 1 + checkEmpty + checkDiff;
                    MoveElement( storySection, startPoint, endPoint, arrayPoint );
                }
                else if ( !( checkTagName[0] != checkTagName2 ) )
                {
                    endPoint = storySectionCounter;
                    MoveElement( storySection, startPoint, endPoint, arrayPoint );
                }
            }
            else if ( checkTagName2Full.search( 'Empty' ) > -1 )
                checkEmpty++;
            else if ( ( checkTagName[0] != checkTagName2 ) || ( checkTagName2.search( 'Title' ) > -1 ) )
            {
                check = 0;
                endPoint = storySectionCounter + 1 + checkEmpty;
                checkEmpty = 0;
                MoveElement( storySection, startPoint, endPoint, arrayPoint );
            }
            else
                checkEmpty = 0;
        } while ( check )
        if ( endPoint < storySection.count() - 1 )
        {
            for ( i = 1 ; ( i ) + endPoint < storySection.count() ; i++ )
            {
                if ( storySection[endPoint + 1].markupTag.name.search( 'Empty' ) > -1 )
                    storySection[endPoint + 1].move( LocationOptions.AT_END, storySection[endPoint] );
                else
                    break;
            }
        }
        Child( storySection[endPoint].xmlElements );
        return storySectionCounter;
    } catch ( err ) { alert( 'MakePA 함수 에러', '오류' ); }
}
// MakePA 에서 호출
// PA 로 묶은 요소들 중에 _가 있는 요소들을 자식 요소로 지정해주는 함수
function Child( PAsection )
{
    try
    {
        var PAcounter = PAsection.count() - 1;
        for ( ; PAcounter > -1 ; PAcounter-- )
        {
            PAname = PAsection[PAcounter].markupTag.name;
            checkPAname = PAname.split( "_" );
            PAlength = checkPAname.length;
            checkPaName0 = checkPAname[PAlength - 1];
            if ( PAlength > 1 )
            {
                if ( !( checkPaName0 != 'Title' ) || !( checkPaName0 != 'Indent' ) )
                    continue;
                startPoint = PAcounter;
                do
                {
                    PAname = PAsection[--PAcounter].markupTag.name;
                    checkPAname = PAname.split( "_" );
                } while ( checkPAname.length > 1 && PAcounter > -1 );
                arrayPoint = PAcounter;
                endPoint = ++PAcounter;
                if ( endPoint )
                {
                    if ( PAsection[endPoint].markupTag.name.search( 'Numbered_Bullet' ) > -1 )
                    {
                        AddElement( PAsection, 'PA_' + PAsection[endPoint].markupTag.name, startPoint );
                        MoveElement( PAsection, startPoint, endPoint, startPoint + 1 );
                        rMoveElement( PAsection, endPoint, endPoint, arrayPoint );
                    }
                    else rMoveElement( PAsection, startPoint, endPoint, arrayPoint );
                }
            }
        }
    } catch ( err ) { alert( 'Child 함수 에러', '오류' ); }
}

// RemakeStruecture 에서 호출
// Topic을 만들어 Heading 단위로 요소를 묶어준다
function MakeTopic( myDoc, myStory, i, HeadingCounter )
{
    var j = i;
    var myTopic = [];
    var myTags = myDoc.xmlTags;
    var myTagsCount = myTags.count() - 1;
    for ( ; myTagsCount > -1 ; myTagsCount-- )
        if ( myTags[myTagsCount].name.search( 'Heading' ) > -1 )
        {
            if ( myTags[myTagsCount].name.search( j ) > -1 )
            {
                myTopic.push( myTags[myTagsCount].name );
                myTopic.sort( SortNumber );
            }
            else if ( !( j != 2 ) )
            {
                if ( myTags[myTagsCount].name.search( 'DIVX' ) > -1 )
                {
                    myTopic.push( myTags[myTagsCount].name );
                    myTopic.sort( SortNumber );
                }
            }
            else if ( !( j != 1 ) )
            {
                if ( myTags[myTagsCount].name.search( 'Trouble' ) > -1 )
                {
                    myTopic.push( myTags[myTagsCount].name );
                    myTopic.sort( SortNumber );
                }
            }
        }

    var topicListCount = null;
    var myTopicList = [];
    try
    {
        var myTopicLength = myTopic.length - 1;
        for ( ; myTopicLength > -1 ; myTopicLength-- )
        {
            try { myTopicList = myTopicList.concat( myStory.itemByName( myTopic[myTopicLength] ).index ); }
            catch ( err ) { continue; }
        }
        myTopicList.sort( SortNumber );
        topicListCount = myTopicList.length - 1;
    }
    catch ( err )
    { topicListCount = -1; }
    var startPoint = myStory.count() - 1;
    for ( ; topicListCount > -1 ; topicListCount-- )
    {
        var endPoint = myTopicList[topicListCount];
        var cTopic = myStory.item( endPoint );
        var stringValue = cTopic.contents;
        AddElement( myStory, 'Topic', stringValue, j, startPoint );
        var arrayPoint = startPoint + 1;
        MoveElement( myStory, startPoint, endPoint, arrayPoint );
        var startPoint = myTopicList[topicListCount] - 1;
        if ( j < HeadingCounter )
            MakeTopic( myDoc, myStory[myTopicList[topicListCount]].xmlElements, j + 1, HeadingCounter );
    }
}

// RemakeStructure 에서 호출
// 타이틀에 해당하는 요소들을 묶어서 Topic_Head 를 만들어 주는 함수
function MakeHead( myDoc )
{
    try
    {
        var myStory = myDoc.xmlElements[0].xmlElements;
        var myStoryCounter = myStory.count();
        for ( var i = 0 ; i < myStoryCounter ; i++ )
        {
            storySection = myStory[i];
            if ( storySection.markupTag.name != 'Story' ) continue;
            var SectionTitleName = storySection.xmlElements[0].markupTag.name;
            if ( !( SectionTitleName != 'SectionTitle' ) || !( SectionTitleName != 'Contents' ) )
            {
                StringValue = storySection.xmlElements[0].contents;
                storySectionCounter = storySection.xmlElements.count();
                AddElement( storySection, StringValue );
                for ( j = 0 ; j < storySectionCounter ; j++ )
                {
                    storySectionFirst = storySection.xmlElements[1];
                    if ( storySectionFirst.markupTag.name != 'Topic' )
                        storySectionFirst.move( LocationOptions.AT_END, storySection.xmlElements[0] );
                    else break;
                }
            }
        }
    } catch ( err ) { alert( 'MakeHead 함수 에러', '오류' ); }
}

// RemakeStructure 에서 호출
// Topic_Head 내에서 Topic을 만들어 주는 함수
function MakeTopicInHead( myDoc )
{
    try
    {
        var rootSection = myDoc.xmlElements[0].xmlElements;
        rootSectionCounter = rootSection.count() - 1;
        for ( ; rootSectionCounter > -1 ; rootSectionCounter-- )
        {
            if ( rootSection[rootSectionCounter].markupTag.name != 'Story' )
                continue;
            headSection = rootSection[rootSectionCounter].xmlElements[0].xmlElements;
            headSectionCounter = headSection.count() - 1;
            startPoint = headSectionCounter;
            for ( ; headSectionCounter > -1 ; headSectionCounter-- )
            {
                TagName = headSection[headSectionCounter].markupTag.name;
                StringValue = headSection[headSectionCounter].contents;
                if ( TagName.search( 'SectionTitle_Sub' ) > -1 )
                {
                    styleName = myDoc.paragraphStyles.itemByName( TagName );
                    endPoint = headSectionCounter;
                    AddElement( headSection, 'Topic', StringValue, i, startPoint );
                    arrayPoint = startPoint + 1;
                    MoveElement( headSection, startPoint, endPoint, arrayPoint );
                    startPoint = headSectionCounter - 1;
                }
            }
        }
    } catch ( err ) { alert( 'MakeTopicInHead 함수 에러', '오류' ); }
}

// Topic과 PA 요소를 만들때 쓰이는 함수
function AddElement( SectionLevel, ElementName, StringValue, Depth, ArrayPoint )
{
    var args_len = arguments.length;
    try
    {
        switch ( args_len )
        {
            case 2:
                SectionLevel.xmlElements.add( { markupTag: 'Topic_Head' } );
                SectionLevel.xmlElements[-1].xmlAttributes.add( 'ID', ElementName );  // ElementName = StringValue
                SectionLevel.xmlElements[-1].move( LocationOptions.AT_BEGINNING, SectionLevel );
                break;
            case 3:
                SectionLevel.add( { markupTag: ElementName } );
                SectionLevel[-1].move( LocationOptions.AFTER, SectionLevel[StringValue] ); // StringValue = ArrayPoint
                break;
            default:
                SectionLevel.add( { markupTag: ElementName } );
                SectionLevel[-1].xmlAttributes.add( 'ID', StringValue );
                SectionLevel[-1].xmlAttributes.add( 'Depth', Depth.toString() );
                SectionLevel[-1].move( LocationOptions.AFTER, SectionLevel[ArrayPoint] );
        }
    } catch ( err ) { alert( 'AddElement 함수 에러', '오류' ); }
}

// RemakeStructure 계열에서 사용
// 요소를 원하는 곳으로 이동시켜준다.
function MoveElement( SectionLevel, StartPoint, EndPoint, ArrayPoint )
{
    try
    {
        var loop = StartPoint - EndPoint;
        do
            SectionLevel[StartPoint--].move( LocationOptions.AT_BEGINNING, SectionLevel[ArrayPoint--] );
        while ( loop-- );
    } catch ( err ) { alert( 'MoveElement 함수 에러', '오류' ); }
}
// RemakeStructure 계열에서 사용
// MoveElement와 시작점과 끝이 반대
function rMoveElement( SectionLevel, StartPoint, EndPoint, ArrayPoint ) {
    try {
        var loop = StartPoint - EndPoint;
        do
            SectionLevel[EndPoint].move( LocationOptions.AT_END, SectionLevel[ArrayPoint] );
        while ( loop-- );
    } catch ( err ) { alert( 'rMoveElement 함수 에러', '오류' ); }
}

// 스타일에 태그를 매핑해준다.
function MappingToStyle( myDoc ) {
    try {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

        var tables = myDoc.xmlElements[0].evaluateXPathExpression("descendant::Table[not(@HeaderRowCount)]");
        var tables_count = tables.length;
        if (tables_count > 0) {
            for ( var i=0 ; i<tables_count ; i++ ) {
                var table = tables[i].tables[0];
                tables[i].xmlAttributes.add('HeaderRowCount', table.headerRowCount + '');
                var rows = table.rows;
                var rows_count = rows.length;
                for ( var j=0 ; j<rows_count ; j++ )
                    rows[j].cells[0].associatedXMLElement.xmlAttributes.add('newrow', 'newrow');
                var columns = table.columns;
                var columns_count = columns.length;
                var widths = [];
                for ( var j=0 ; j<columns_count ; j++ )
                    widths = widths.concat((columns[j].width/table.width).toFixed(2) + '*');
                tables[i].xmlAttributes.add('colspecs', widths.join(':'));
            }
    
            var cells = myDoc.xmlElements[0].evaluateXPathExpression("descendant::Cell[not(@namest)]");
            var cells_count = cells.length;
            for ( var i=0 ; i<cells_count ; i++) {
                var cell = cells[i].cells[0];
                if ( cell.columnSpan > 1 ) {
                    var namest = cell.parentColumn.index + 1;
                    var nameend = namest + cell.columnSpan - 1;
                    cells[i].xmlAttributes.add('namest', 'col' + namest);
                    cells[i].xmlAttributes.add('nameend', 'col' + nameend);
                }
            }
        }
    
        var need_id = myDoc.xmlElements[0].evaluateXPathExpression("descendant::*[starts-with(name(), 'Chapter') or starts-with(name(), 'Heading') or name()='Description-NoHTML'][not(@id)]");
        var needs_count = need_id.length;
        for ( var i=0 ; i<needs_count ; i++ )
            need_id[i].xmlAttributes.add('id', 'd' + need_id[i].id);
            
        var nested = myDoc.xmlElements[0].evaluateXPathExpression("descendant::Empty[not(@nested)][preceding-sibling::*[name()!='Empty'][1][starts-with(name(), 'OrderList') or starts-with(name(), 'UnorderList')]]");
        var nested_count = nested.length;
        for ( var i=0 ; i<nested_count ; i++ )
            nested[i].xmlAttributes.add('nested', 'nested');
    
        var TagList = myDoc.xmlTags.everyItem();
        var TagCounter = myDoc.xmlTags.count();
        var None = myDoc.characterStyles[0];
        var xmlMaps = myDoc.xmlImportMaps;

        for ( var i = 0 ; i < TagCounter ; i++ ) {
            // Cell, Story, Table, Topic*, PA_* 들은 따로 처리 해준다.
            TagName = TagList.name[i];
            if ( ( ( TagName.search( 'PA_' ) < 0 ) && ( TagName.search( 'Topic' ) < 0 ) ) && ( ( TagName != 'Cell' ) && ( TagName != 'Story' ) && ( TagName != 'Table' ) ) )
                xmlMaps.add( TagName, TagName );
            else
                xmlMaps.add( TagName, None );
        }
        myDoc.mapXMLTagsToStyles();
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    } catch (ex) {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
        alert( 'MappingToStyle함수 에러(' + myDoc.name + '):' + ex, '오류' );
        errorCount++;
    }
}

// 본문에서 호출
// 구조에서 문자 스타일의 오류를 체크한 후에 로그 작성.
function MakeCharErrorLog( myDoc, myXmlElements, count, MMIcondition )
{
    try
    {
        var ElementsList = count;
        myCharacterStyle = myDoc.characterStyles;
        var checkC_Image = 0; var check_Style = null; var tag_Name = null;
        for ( var i = 0 ; i < ElementsList ; i++ )
        {
            myXmlElementItem = myXmlElements[i];
            myXmlElementItemElements = myXmlElements[i].xmlElements;
            tag_Name = myXmlElementItem.markupTag.name;
            if ( MMIcondition && !( ( tag_Name != 'MMI' ) && ( tag_Name != 'MMI_NoBold' ) ) )
            {
                IDcounter = ++MMIcounter;
                var myChapterName = myDoc.name.split( '.' )[0];
                try { app.activeDocument; IDcounter = myChapterName + '_' + IDcounter.toString(); }
                catch ( err ) { IDcounter = IDcounter.toString(); }
                myXmlElementItem.xmlAttributes.add( 'ID', IDcounter );
                myXmlElementItem.xmlAttributes.add( 'Chapter', myChapterName );
                try { myXmlElementItem.xmlAttributes.add( 'Page', myXmlElementItem.paragraphs[0].parentTextFrames[0].parentPage.name ); }
                catch ( err ) { myXmlElementItem.xmlAttributes.add( 'Page', myXmlElementItem.paragraphs[0].parentTextFrames[0].parent.name ); }
            }
            var check = 0;
            checkC_Image = 0;
            try
            {
                //tag_Name = myXmlElementItem.markupTag.name;
                if ( ( tag_Name.search( 'Image' ) > -1 ) || ( tag_Name.search( 'Icon' ) > -1 ) || ( tag_Name.search( 'Story' ) > -1 ) || ( tag_Name.search( 'Img' ) > -1 ) )
                    checkC_Image = -1;
                check_Style = myCharacterStyle.itemByName( tag_Name ).name;
                check = 1;
            }
            catch ( err ) { check = 0; }
            if ( check )
            {
                contentsValue = myXmlElementItem.contents;
                if ( myXmlElementItem.parent.markupTag.name.search( 'Empty' ) > -1 && tag_Name.search( 'Story' ) < 0 )
                    MakeErrorAddress( myXmlElementItem, '\r\nError: Empty 스타일 다음에 문자스타일이 올 수 없습니다' );
                else if ( !( contentsValue[0] != ' ' ) || !( contentsValue[contentsValue.length - 1] != ' ' ) )
                {
                    if ( ( check_Style.search( 'C_CircleNumber' ) > -1 ) && !( contentsValue.length != 3 ) )
                    { }
                    else
                    {
                        MakeErrorAddress( myXmlElementItem, '\r\nError: 문자스타일에 불필요한 공백이 추가되었습니다' );
                        /*if ( checkC_Image<0)
                        {
                            try {
                                myXmlElementItem.xmlElements[0].xmlAttributes[0].value;
                                check_CharacterError += 'Error: 이미지의 문자스타일이 잘못되었습니다' + contentsValue+'\r\n\r\n';
                                write_ErrorAddress += index + '\r\n';
                            } catch(err) {}
                        }*/
                    }
                }
                else if ( contentsValue.search( '→' ) > -1  || contentsValue.search( '←' ) > -1)
                {
                    if ( contentsValue.length > 1 )
                        MakeErrorAddress( myXmlElementItem, '\r\nError: C_Singlestep에 다른 내용이 들어가 있습니다' );
                    else if ( check_Style.search( 'C_Singlestep' ) < 0 && !integrate )
                        MakeErrorAddress( myXmlElementItem, '\r\nError: C_Singlestep 문자스타일을 적용해 주세요' );
                    else if ( check_Style.search( 'C_SingleStep' ) < 0 && integrate )
                        MakeErrorAddress( myXmlElementItem, '\r\nError: C_SingleStep 문자스타일을 적용해 주세요' );
                }
                else if ( contentsValue.search( '►' ) > -1 || contentsValue.search( '◄' ) > -1 )
                {
                    if ( contentsValue.length > 1 )
                        MakeErrorAddress( myXmlElementItem, '\r\nError: 문자스타일에 다른 내용이 들어가 있습니다' );
                    if ( ( check_Style.search( 'C_Seepage' ) < 0 ) && ( check_Style.search( 'C_Symbol' ) < 0 ) && ( check_Style.search( 'C_Arial_R' ) < 0 ) && ( check_Style.search( 'C_SingleStep' ) < 0 ) )
                        MakeErrorAddress( myXmlElementItem, '\r\nError: C_Seepage 나 C_Symbol 문자스타일을 적용해 주세요' );
                }
                else if ( myXmlElementItem.words.length < 1 )
                {
                    if ( myXmlElementItem.characters.count() < 2 )
                    {
                        if ( checkC_Image < 0 )
                        {
                            if ( !integrate )
                                MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지의 문자스타일이 잘못되었습니다' );
                            else
                            {
                                if ( myXmlElementItem.evaluateXPathExpression('descendant-or-self::*[@href]').length < 1 )
                                     MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' );
                            /*
                                    try { if ( myXmlElementItem.xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                                    catch ( err )
                                    {
                                        try { if ( myXmlElementItem.xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                                        catch ( err )
                                        {
                                            try { if ( myXmlElementItem.xmlElements[0].xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                                            catch ( err ) { MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' ); }
                                        }
                                    }
                                    */
                            }
                        }
                        else if ( !( contentsValue.charCodeAt( 0 ) != 13 ) )
                            MakeErrorAddress( myXmlElementItem, '\r\nError: 단락과 단락 사이에 태그가 사용되었습니다' );
                        else if ( !( contentsValue.charCodeAt( 0 ) != 65279 ) )
                            MakeErrorAddress( myXmlElementItem, '\r\nError: 불필요한 문자스타일이 사용되었습니다' );
                        else if (contentsValue != '则'
                                && contentsValue != '無'
                                && contentsValue != '點'
                                && contentsValue != '圈'
                                && contentsValue != '則'
                                && contentsValue != '앱'
                                && contentsValue != '방'
                                && contentsValue != '홈'
                                && contentsValue != '팁'
                                && contentsValue != '가'
                                && contentsValue != '예'
                                && contentsValue != '펜'
                                )
                            MakeErrorAddress( myXmlElementItem, '\r\nError: 문자스타일 공백 오류가 발생하였습니다' );
                    }
                    else if ( contentsValue.length < 1 )
                        MakeErrorAddress( myXmlElementItem, '\r\nError: 4' );
                }
                else if ( !( contentsValue.length != 1 ) && !( contentsValue.charCodeAt( 0 ) != 65279 ) )
                    MakeErrorAddress( myXmlElementItem, '\r\nError: 불필요한 문자스타일이 사용되었습니다' );
                else if ( checkC_Image < 0 )
                {
                    if ( myXmlElementItem.evaluateXPathExpression('descendant-or-self::*[@href]').length < 1 )
                         MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' );
                  /*
                        try { if ( myXmlElementItem.xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                        catch ( err )
                        {
                            try { if ( myXmlElementItem.xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                            catch ( err )
                            {
                                try { if ( myXmlElementItem.xmlElements[0].xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                                catch ( err ) { MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' ); }
                            }
                        }
                        */
                }
            }
            else if ( MMIcondition && !( tag_Name != 'Table' ) )
            {
                if ( myXmlElementItem.parent.markupTag.name.search( 'Empty' ) < 0 )
                    MakeErrorAddress( myXmlElementItem, '\r\nError: Table 스타일이 Empty가 아닌 다른 상위값을 가지고 있습니다' );
            }
            else if ( checkC_Image < 0 && tag_Name != 'Story' )
            {
                if ( myXmlElementItem.evaluateXPathExpression('descendant-or-self::*[@href]').length < 1 )
                     MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' );
                /*
                    try { if ( myXmlElementItem.xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                    catch ( err )
                    {
                        try { if ( myXmlElementItem.xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                        catch ( err )
                        {
                            try { if ( myXmlElementItem.xmlElements[0].xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                            catch ( err ) { MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' ); }
                        }
                    }
                    */
            }
            check = 0;
            try
            {
                check_Style = myDoc.paragraphStyles.itemByName( tag_Name ).name;
                check = 1;
                contentsValue = myXmlElementItem.contents;
                if ( ( check_Style.search( 'Image' ) > -1 ) || ( check_Style.search( 'Icon' ) > -1 ) || ( check_Style.search( 'Story' ) > -1 ) || ( check_Style.search( 'Img' ) > -1 ) )
                    checkC_Image = -1;
            }
            catch ( err )
            { check = 0; }
            if ( check )
            {
                //if ( !(contentsValue[0] != ' ') || !(contentsValue[contentsValue.length-1] != ' ') )
                //MakeErrorAddress(myXmlElementItem, '\r\nError: 단락스타일에 불필요한 공백이 추가되었습니다');
                if ( checkC_Image < 0 )
                {
                  if ( myXmlElementItem.evaluateXPathExpression('descendant-or-self::*[@href]').length < 1 )
                     MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' );
                  /*
                        try { if ( myXmlElementItem.xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                        catch ( err )
                        {
                            try { if ( myXmlElementItem.xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                            catch ( err )
                            {
                                try { if ( myXmlElementItem.xmlElements[0].xmlElements[0].xmlAttributes[0].name != 'href' ) MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지에 링크가 없습니다' ); }
                                catch ( err ) { MakeErrorAddress( myXmlElementItem, '\r\nError: 이미지가 아닌 곳에 이미지 문자스타일이 들어갔거나 링크가 없습니다' ); }
                            }
                        }
                        */
                }
                else if ( check_Style.search( 'Empty' ) > -1 )
                {
                    try
                    {
                        //myXmlElementItem.xmlAttributes[0].value;
                        if (myXmlElementItem.evaluateXPathExpression("self::*[@*[name()!='id']]").length > 0)
                            MakeErrorAddress( myXmlElementItem, '\r\nError: Empty 스타일에 이미지가 쓰였습니다' );
                        else
                        {
                            if ( contentsValue.length > 2 )
                            MakeErrorAddress( myXmlElementItem, '\r\nError: Empty 스타일에 다른 내용이 들어왔습니다' );
                            else if ( ( contentsValue[0].charCodeAt( 0 ) != 65532 ) && ( contentsValue[0].charCodeAt( 0 ) != 22 ) && ( contentsValue[0].charCodeAt( 0 ) != 13 ) )
                                MakeErrorAddress( myXmlElementItem, '\r\nError: Empty 스타일에 다른 내용이 들어왔습니다' );
                        }
                    }
                    catch ( err )
                    {/*
                            if ( contentsValue.length > 2 )
                                MakeErrorAddress( myXmlElementItem, '\r\nError: Empty 스타일에 다른 내용이 들어왔습니다' );
                            else if ( ( contentsValue[0].charCodeAt( 0 ) != 65532 ) && ( contentsValue[0].charCodeAt( 0 ) != 22 ) && ( contentsValue[0].charCodeAt( 0 ) != 13 ) )
                                MakeErrorAddress( myXmlElementItem, '\r\nError: Empty 스타일에 다른 내용이 들어왔습니다' );
                        */}
                }
                //else if ( myXmlElementItem.markupTag.name == "Blank" )
                //{
                //var reg = /\S+/;
                //if ( reg.test(contentsValue) || contentsValue.length > 1)
                //MakeErrorAddress(myXmlElementItem, '\r\nError: Blank 스타일에 내용이 들어갈 수 없습니다');
                //}
                else if ( myXmlElementItem.words.length < 1 )
                {
                    if ( myXmlElementItem.characters.count() < 2 )
                    {
                        if ( !( contentsValue.charCodeAt( 0 ) != 13 ) && myXmlElementItem.markupTag.name != 'Empty' )
                        {
                            if (!myXmlElementItem.paragraphs[0].appliedParagraphStyle.ruleAbove && !myXmlElementItem.paragraphs[0].appliedParagraphStyle.ruleBelow)
                                MakeErrorAddress( myXmlElementItem, '\r\nError: 줄넘김에 태그가 사용되었습니다' );
                        }
                        else if ( contentsValue.charCodeAt( 0 ) != 22 )
                        {
                            if ( check_Style != 'Note' && check_Style != 'CellBody' )
                                MakeErrorAddress( myXmlElementItem, '\r\nError: 빈 공간에 태그가 사용되었습니다' );
                            else if ( myXmlElementItem.pdfs.count() < 1 ) MakeErrorAddress( myXmlElementItem, '\r\nError: 빈 공간에 태그가 사용되었습니다' );
                        }
                    }
                    else if ( !( contentsValue.length != 1 ) && !( contentsValue.charCodeAt( 0 ) != 65279 ) )
                        MakeErrorAddress( myXmlElementItem, '\r\nError: 불필요한 스타일이 사용되었습니다' );
                }
            }
            var recount = myXmlElementItemElements.count();
            if ( recount )
                MakeCharErrorLog( myDoc, myXmlElementItemElements, recount, MMIcondition );
        }
    }
    catch ( err ) { alert( 'MakeCharErrorLog 함수 에러', '오류' ); return -1; }
    return 1;
}

//오류가 난 곳의 주소를 변수로 저장
function MakeErrorAddress( myXmlElementItem, errorStyle )
{
    xmlParentItem = myXmlElementItem.parent;
    var contentsValue = '';
    if ( ( xmlParentItem.markupTag.name != 'Story' ) && ( xmlParentItem.markupTag.name != 'Empty' ) && ( xmlParentItem.markupTag.name != 'Cell' ) )
        contentsValue = '\r\n전체문장: ' + xmlParentItem.contents;
    //else if ( !(myXmlElementItem.markupTag.name != 'C_Image') )

    contentsValue = errorStyle + contentsValue;
    try { myXmlElementItem.select( SelectionOptions.ADD_TO ); } catch ( err ) { };
    var index = [];
    j = 0;
    var test = myXmlElementItem;
    do
    {
        index[j++] = test.index;
        test = test.parent;
    } while ( test.markupTag.name != 'Root' )
    index.reverse();
    check_CharacterError += ++errorCount + '.Address: ' + index + contentsValue + '\r\n\r\n';
    write_ErrorAddress += index + '\r\n';
}

//오류가 난 곳의 주소를 수동으로 입력해서 검색
function FindErrorPoint()
{
    var myRoot = app.activeDocument.xmlElements.item( 0 )

    errorAddress = [InputPointDialog()];
    if ( errorAddress < 0 ) return -1;
    errorAddress = errorAddress[0].split( ',' ).reverse();

    var errorCount = errorAddress.length - 1;
    var errorPoint = ErrorPointSelect( myRoot, errorCount );

    try { errorPoint.select(); }
    catch ( err ) { alert( '잘못된 입력입니다. 주소값을 확인해주세요', '주소입력 오류' ); }
    return 1;
}

//수동으로 입력된 주소값을 XML주소값으로 변환
function ErrorPointSelect( mySection, errorCount )
{
    if ( errorCount > -1 )
    {
        var mySection = null;
        try { mySection = mySection.xmlElements[errorAddress[errorCount]]; }
        catch ( err ) { return -1; }
        errorCount--;
        return ErrorPointSelect( mySection, errorCount );
    }
    else
        return mySection;
}

//오류가 난 곳의 주소값을 수동으로 입력
function InputPointDialog()
{
    var SearchPoint = app.dialogs.add( { name: '주소로 오류 찾기' } );
    var InputArea = SearchPoint.dialogColumns.add();
    var InputPoint = InputArea.dialogRows.add();
    var Caption = InputArea.dialogRows.add();
    InputPoint.staticTexts.add( { staticLabel: '주소 입력  (예: 0,0,0):', minWidth: 150 } );
    Caption.staticTexts.add( { staticLabel: 'Log파일을 참조해서 주소를 입력하세요', alignment: 'right' } );
    var PointValue = InputPoint.textEditboxes.add( { minWidth: 80, editContents: '0' } );
    if ( SearchPoint.show() )
        return PointValue.editContents;
    else
        return -1;
}

// 이미지 크기를 일괄적으로 조정할때 그 크기값을 입력받는 함수
function InputImageSize()
{
    do
    {
        var SearchPoint = app.dialogs.add( { name: 'Image Size' } );
        var InputArea = SearchPoint.dialogColumns.add();
        var InputPoint = InputArea.dialogRows.add();
        InputPoint.staticTexts.add( { staticLabel: 'Input Size (170%일 경우 170):', minWidth: 100 } );
        var PointValue = InputPoint.textEditboxes.add( { minWidth: 100, editContents: '0' } );
        if ( SearchPoint.show() - 1 )
            return -1;
        else if ( +PointValue.editContents <= 0 )
            alert( '잘못된 입력입니다\r\n0보다 큰 수를 입력해주세요', '입력오류' );
    } while ( +PointValue.editContents <= 0 )
    return PointValue.editContents;
}

// 이미지들의 링크 폴더를 재설정 해주는 함수. 설정전에 문서의 종류나 숫자등을 정해주는 함수
function RelinkImage()
{
    resultLog = '링크' + resultLog;
    var check = 0;
    var docNames = [];
    var returnValue = Select_Contents( true );
    if ( returnValue < 0 )
        return -1;
    else if ( returnValue == 0 || returnValue == 'continue' )
        return 0;
    docNames = returnValue;
    var imagePath = InputFolder();
    if ( imagePath < 0 )
        return -1;
    else imagePath += '/'
    if ( onlyDocument )
        LinkImage(imagePath, app.activeDocument);
    else
    {
        var bookDocuments = [];
        var bookDocumentsLength = docNames.length;
        for ( var i = 0 ; i < bookDocumentsLength ; i++ )
        {
            bookDocuments[i] = app.open( docNames[i], false );
            LinkImage( imagePath, bookDocuments[i] );
            try { bookDocuments[i].close( SaveOptions.YES, docNames[i] ); }
            catch ( err )
            {
                errorCount++;
                resultLog += myDoc.name + '실패\r\n';
                alert( docNames[i].name + '의 저장에 실패했습니다' );
            }
        }
        app.activeBook.save();
    }

    if ( errorCount )
        alert( resultLog, '링크폴더 재설정 결과' );
    else
        alert( '링크폴더 재설정이 성공적으로 끝났습니다', '완료' );
    return 1;
}

// 이미지들의 링크 폴더를 실제로 수정해주는 함수
function LinkImage( imagePath, myDoc )
{
    errorCount = 0;
    try
    {
        var allImages = myDoc.allGraphics;
        var imageCount = allImages.length;
        for ( var i = 0 ; i < imageCount ; i++ )
        {
            try { linkName = allImages[i].itemLink.name; }
            catch ( err ) { continue; }
            fullPath = File( imagePath + linkName );
            try { allImages[i].itemLink.relink( fullPath ); }
            catch ( err ) { errorLog += linkName + ' '; errorCount++; }
        }
        imaeCount = myDoc.allGraphics.length;
        for ( i = 0 ; i < imageCount ; i++ )
        {
            try { allImages[i].itemLink.update(); }
            catch ( err ) { }
        }
    } catch ( err ) { alert( 'LinkImage 함수 에러', '오류' ); }
    if ( errorCount )
        resultLog += myDoc.name + ' 안에 업데이트되지 않은 이미지가 있습니다.\r\n';
}

//영문 색인 목록을 대상 색인 목록으로 호출해오는 함수
function ImportIndex()
{
    var check = 0;
    var errorCondition = 0;
    var checkError = 0;
    var checkBookType = 0;
    var books = null;
    var XMLToolStatus = null;
    var myDoc = null;
    var book = null;
    var booksCount = null;
    var docNames = [];
    var sourceFile = 0;
    var indexResult = null;

    do
    {
        checkBookType = 0;
        try { myDoc = app.activeDocument; check = 1; } catch ( err ) { }
        if ( !check )
        {
            try
            {
                books = app.books;
                booksCount = app.books.count();
                if ( booksCount > 1 )
                    var book = SelectBook( books );
                else
                    book = app.activeBook;
                var bookContents = book.bookContents;
                try { var bookContentsCount = bookContents.count(); }
                catch ( err )
                {
                    if ( !book )
                        return 0;
                    else
                        return -1;
                }
                for ( var i = 0 ; i < bookContentsCount ; i++ )
                    docNames[i] = bookContents[i];
                docNames = SelectINDD( docNames );
                if ( !docNames )
                    return 0;
                else if ( docNames < 0 )
                    return -1;
                check = 2;
            }
            catch ( err )
            {
                alert( '열려있는 문서나 북이 없습니다', '오류' );
                return -1;
            }
        }
    } while ( checkBookType )
    if ( check == 1 )
    {
        try { myDoc.filePath; }
        catch ( err )
        {
            alert( myDoc.name + '가 한 번도 저장되지 않았습니다.\r\n저장하신 후에 진행해주세요', '오류' );
            return -1;
        }
    }
    if ( check > 1 )
    {
        var res =
            "dialog { alignChildren: 'fill', \
                Success: Panel { orientation: 'row', alignChildren:'center', \
                    text: '문서들이 확인창 없이 자동저장 됩니다', \
                    s: StaticText { text: '진행하시겠습니까?' },\
                }, \
                buttons: Group { orientation: 'row', alignment: 'right', \
                    okBtn: Button { text:'OK', properties:{name:'ok'} }, \
                    cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                } \
            }";
        var result = new Window( res, '문서 자동저장 경고' );
        result.center();
        var select = result.show() - 1;
        if ( select )
            return -1;
        else if ( check > 0 )
        {
            sourceFile = SelectSourceINDD( docNames[0] );
            if ( sourceFile < 0 )
                return -1;
        }
        var bookDocument;
        var bookDocumentsLength = docNames.length;
        for ( var i = 0 ; i < bookDocumentsLength ; i++ )
        {
            bookDocument = app.open( docNames[i], false );
            if ( !bookDocument.xmlElements[0].xmlElements.count() )
            {
                checkError = TaggingForIndex( bookDocument );
                if ( !checkError )
                {
                    indexingLog += bookDocument.name + ' 의 XML구조에 오류가 있습니다\r\n';
                    totalError++;
                    continue;
                }
            }
            sourceFile = File( sourceFile.path + '/' + bookDocument.name );
            errorCondition = MakeIndex( bookDocument, sourceFile );
            if ( errorCondition )
            {
                indexingLog += '\r\n' + errorCondition;
                totalError++;
            }
            bookDocument.close( SaveOptions.YES, docNames[i] );
        }
        app.activeBook.save();
    }
    else if ( check )
    {
        if ( !myDoc.xmlElements[0].xmlElements.count() )
        {
            checkError = TaggingForIndex( myDoc );
            if ( !checkError )
            {
                indexingLog += myDoc.name + ' 의 XML구조에 오류가 있습니다\r\n';
                totalError++;
            }
        }
        sourceFile = SelectSourceINDD( myDoc );
        errorCondition = MakeIndex( myDoc, sourceFile );
        if ( errorCondition )
        {
            indexingLog += '\r\n' + errorCondition;
            totalError++;
        }
    }
    if ( totalError )
        alert( indexingLog, '색인 작업 결과' );
    else
        alert( '색인 작업이 성공적으로 끝났습니다', '색인 작업 결과' );
    return 1;
}

//색인 작업 이전에 태그를 입혀주는 작업
function TaggingForIndex( myDoc ) {
    myDoc.xmlPreferences.defaultCellTagName = "Cell";
    myDoc.xmlPreferences.defaultImageTagName = "C_Image";
    myDoc.xmlPreferences.defaultStoryTagName = "Story";
    myDoc.xmlPreferences.defaultTableTagName = "Table";
    var check = 0;
    var myGraphics = myDoc.allGraphics;
    var myGraphicsCount = myGraphics.length;
    myDoc.deleteUnusedTags();
    myDoc.xmlElements[0].remove();
    check = CheckTags( myDoc, 1, 1 );
    if ( check < 0 )
        return 0;
    ChangeTagColor( myDoc );
    myDoc.mapStylesToXMLTags();  // 구조 등록
    myDoc.mapStylesToXMLTags();  // 이미지를 위해 한번 더 구조 등록
    for ( i = 0 ; i < myGraphicsCount ; i++ ) {
        try { myGraphics[i].autoTag(); }
        catch ( err ) { continue; }
    }
    //specials_to_pi(myDoc);
    Get_HyperlinkInform( myDoc );
    MakeCharErrorLog( myDoc, myDoc.xmlElements[0].xmlElements, myDoc.xmlElements[0].xmlElements.count() );
    MappingToStyle( myDoc );
    //pi_to_breaks_and_specials(myDoc);
    if ( errorCount )
        return 0;
    AddImageSize( myDoc );
    return 1;
}

//이미지를 링크할 폴더를 선택
function InputFolder() {
    var myFolder = new Window( 'dialog' );
    myFolder.center();
    myFolder = Folder.selectDialog( '이미지를 링크시킬 폴더를 선택해주세요' );
    if ( myFolder )
        return myFolder;
    else
        return -1;
}

//
function ChangeInterface( before, after )
{
    var check = 0;
    var docNames = [];
    var myDoc = null;
    try
    {
        myDoc = app.activeDocument;
        check = 1;
    } catch ( err ) { }
    if ( !check )
    {
        docNames = SelectTarget( true );
        if ( !docNames )
            return 0;
        else if ( docNames < 0 )
            return -1;
        else
            check = 2;
    }
    else
    {
        try { myDoc.filePath; }
        catch ( err )
        {
            alert( myDoc.name + '가 한 번도 저장되지 않았습니다.\r\n저장하신 후에 진행해주세요', '오류' );
            return -1;
        }
    }

    if ( check > 1 )
    {
        var res =
            "dialog { alignChildren: 'fill', \
                Success: Panel { orientation: 'row', alignChildren:'center', \
                    text: '문서들이 확인창 없이 자동저장 됩니다', \
                    s: StaticText { text: '진행하시겠습니까?' },\
                }, \
                buttons: Group { orientation: 'row', alignment: 'right', \
                    okBtn: Button { text:'OK', properties:{name:'ok'} }, \
                    cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                } \
            }";
        var result = new Window( res, '문서 자동저장 경고' );
        result.center();
        var select = result.show() - 1;
        if ( select )
            return -1;

        var bookDocumentsLength = docNames.length;
        for ( var i = 0 ; i < bookDocumentsLength ; i++ )
        {
            myDoc = app.open( docNames[i], false );
            try
            {
                var before_name = before;
                var after_name = after;
                var beforestyle = myDoc.characterStyles.itemByName( before_name );
                var afterstyle = myDoc.characterStyles.itemByName( after_name );
                var chn = 'CHN_' + before_name;
                if (beforestyle == null)
                {
                    before_name = 'CHN_' + before;
                    beforestyle = myDoc.characterStyles.itemByName( before_name);
                    chn = null;
                }
                if (beforestyle == null)
                {
                    if (myDoc.name.search('_Cover.') < 0 && myDoc.name.search('_TOC.') < 0 && myDoc.name.search('_Copyright.') < 0 )
                        alert (myDoc.name + ' 문서에 MMI 스타일이 없습니다', '오류');
                    myDoc.close(SaveOptions.NO);
                    continue;
                }
                
                replaceStyle(myDoc, beforestyle, afterstyle, before_name, after_name, chn);
            }
            catch ( err )
            { 
                alert (err);
                myDoc.close(SaveOptions.NO);
                return -1;
            }
            try { myDoc.close( SaveOptions.YES, docNames[i] ); }
            catch ( err )
            {
                errorCount++;
                resultLog += myDoc.name + '실패\r\n';
                alert( myDoc.name + '의 저장에 실패했습니다' );
            }
        }
        app.activeBook.save();
    }
    else if ( check > 0 )
    {
        try
        {
            var beforestyle = myDoc.characterStyles.itemByName( before );
            var afterstyle = myDoc.characterStyles.itemByName( after );
            var chn = 'CHN_' + before;
            if (beforestyle == null)
            {
                before = 'CHN_' + before;
                beforestyle = myDoc.characterStyles.itemByName( before);
                chn = null;
            }
            if (beforestyle == null)
            {
                alert ('MMI 스타일이 없습니다', '오류');
                return -1;
            }
            
            replaceStyle(myDoc, beforestyle, afterstyle, before, after, chn);
        }
        catch ( err )
        {
            alert(err);
            return -1;
        };
    }
    if ( errorCount )
        alert( resultLog, '문서 저장 결과' );
    else
        alert( '스타일 값 변환이 성공적으로 완료되었습니다', '완료' );
    return 1;
}

function replaceStyle(myDoc, beforestyle, afterstyle, before, after, chn)
{
    if (afterstyle != null)
    {
      /*
        afterstyle.remove(beforestyle);
        if (after == 'Iword')
            myDoc.characterStyles.itemByName(after + '-NoBold').remove(myDoc.characterStyles.itemByName(before + '_NoBold'));
        else
            myDoc.characterStyles.itemByName(after + '_NoBold').remove(myDoc.characterStyles.itemByName(before + '-NoBold'));
         */
        beforestyle.remove (afterstyle);
        if (after == 'Iword')
            myDoc.characterStyles.itemByName(before + '_NoBold').remove(myDoc.characterStyles.itemByName(after + '-NoBold'));
        else
            myDoc.characterStyles.itemByName(before + '-NoBold').remove(myDoc.characterStyles.itemByName(after + '_NoBold'));
    }
    else
    {
        myDoc.characterStyles.itemByName( before ).name = after;
        if ( before == 'Iword' )
            myDoc.characterStyles.itemByName( before + '-NoBold' ).name = after + '_NoBold';
        else
            myDoc.characterStyles.itemByName( before + '_NoBold' ).name = after + '-NoBold';
    }
    
    if (before != 'Iword' && chn != null && myDoc.characterStyles.itemByName(chn) != null)
    {
        myDoc.characterStyles.itemByName(chn).remove (myDoc.characterStyles.itemByName(after));
        myDoc.characterStyles.itemByName(chn + '_NoBold').remove(myDoc.characterStyles.itemByName(after + '-NoBold'));
    }
    
}

// INDD파일을 IDML이나 INX로 내보내는 함수
function ExportIDML()
{
    var check = 0;
    var myFolder = 0;
    var myDoc = null;
    var idmlFile = null;
    var docNames = [];

    try { myDoc = app.activeDocument; check = 1; } catch ( err ) { }
    if ( !check )
    {
        if (export_all)
            docNames = SelectTarget( true );
        else
            docNames = SelectTarget( false );
        if ( !docNames )
            return 0;
        else if ( docNames < 0 )
            return -1;
        else
            check = 2;
    }
    myFolder = Folder.selectDialog( 'Select your specific folder' );
    if ( !myFolder )
        return 0;
    if ( check < 2 )
    {
        progresspanel.progressbar.value = 0;
        progresspanel.progressbar.maxvalue = 1;
        progresspanel.show();
        Checker( myDoc, false );
        DeleteTag( myDoc );
        if ( parseInt( app.version ) > 6 )
        {
            idmlFile = new File( myFolder + '/' + myDoc.name.split( '.' )[0] + '.idml' );
            myDoc.exportFile( ExportFormat.INDESIGN_MARKUP, idmlFile, true );
        }
        else
        {
            idmlFile = new File( myFolder + '/' + myDoc.name.split( '.' )[0] + '.inx' );
            myDoc.exportFile( ExportFormat.INDESIGN_INTERCHANGE, idmlFile, true );
        }
        progresspanel.progressbar.value = 1;
    }
    else
    {
        var bookDocument;
        var bookDocumentsLength = docNames.length;
        progresspanel.progressbar.maxvalue = bookDocumentsLength;
        progresspanel.show();
        for ( var i = 0 ; i < bookDocumentsLength; i++ )
        {
            progresspanel.progressbar.value = i + 1;
            if ( export_all || (( docNames[i].name.split( '.' )[0].search( 'Cover' ) < 0 ) && ( docNames[i].name.split( '.' )[0].search( 'TOC' ) < 0 )) )
            {
                bookDocument = app.open( docNames[i], false );
                Checker( bookDocument, false );
                DeleteTag( bookDocument );
                if ( parseInt( app.version ) > 6 )
                {
                    idmlFile = new File( myFolder + '/' + bookDocument.name.split( '.' )[0] + '.idml' );
                    bookDocument.exportFile( ExportFormat.INDESIGN_MARKUP, idmlFile );
                }
                else
                {
                    idmlFile = new File( myFolder + '/' + bookDocument.name.split( '.' )[0] + '.inx' );
                    bookDocument.exportFile( ExportFormat.INDESIGN_INTERCHANGE, idmlFile );
                }
                bookDocument.close( SaveOptions.NO );
            }
        }
    }
    progresspanel.hide();
    alert( '완료되었습니다', '완료' );
}

////////////////////////////////////////////문자열을 16진수로 에스케이프 해서, 「\x{hex}」라고 하는 형태로 되돌려줌
function my_escape( str )
{
    tmp_str = escape( str );
    return tmp_str.replace( /\%u([0-9A-F]+)/i, "\\x{$1}" )
}
//////////////////////////////////////////// 정규표현검색으로 카타카나를 16진수로 변환
function katakana2hex()
{
    var find_str = app.findGrepPreferences.findWhat;
    while ( /([ァ-ヴ])/.exec( find_str ) )
        find_str = find_str.replace( /([ァ-ヴ])/, my_escape( RegExp.$1 ) );
    app.findGrepPreferences.findWhat = find_str;    //검색문자설정
}
////////////////////////////////////////////에러처리
function MyErrorExit( mess )
{
    if ( arguments.length > 0 )
        alert( mess, '작동 중지' );
    exit();
}

////////////////////////////////////////////이하 본문
function RedefineStyles( my_Queries, myDoc, del_Image )
{
    errorCount = 0;
    var reDefineLog = "";
    //var docs = app.documents;
    //var len = docs.length;
    //if (len<1) MyErrorExit("문서가 열려 있지 않습니다.");

    //for(var index=0; index<len; index++)
    //{
    //var mydocument = docs[index];
    var mydocument = myDoc;
    var docName = mydocument.name;
    //alert(docName, '대상 문서입니다');
    //검색범위를 지정
    var my_range_obj = mydocument;

    if ( my_range_obj.characterStyles.itemByName( 'C_Image' ) == null && integrate && !del_Image )
    {
        my_range_obj.characterStyles.add( { name: 'C_Image' } );
        if ( !no_intergrate )
            my_range_obj.characterStyles.itemByName( 'C_Image' ).baselineShift = -2;
    }

    //검색치환
    //reDefineLog = '대상문서 : ' + docName.split('.')[0] + '\r\n';
    for ( var i = 0 ; i < my_Queries.length ; i++ )
    {
        try
        {
            var my_mode = my_Queries[i][0];
            var my_query_name = my_Queries[i][1];
            var my_Style = my_Queries[i][2];
            if ( ( my_range_obj.characterStyles.itemByName( my_Style ) == null ) && ( my_range_obj.paragraphStyles.itemByName( my_Style ) == null ) )
            { }   //reDefineLog += '\r\n ' + my_query_name + ': 해당 스타일이 없습니다(' + my_Style + ')';
            else if ( my_mode == "text" )
            {
                my_mode = SearchModes.TEXT_SEARCH;
                app.loadFindChangeQuery( my_query_name, my_mode );    //쿼리를 세트
                my_range_obj.changeText();
                //reDefineLog += '\r\n ' + my_query_name + ': ' + my_range_obj.changeText().length + ' 개';
            }
            else if ( my_mode == "grep" )
            {
                my_mode = SearchModes.GREP_SEARCH;
                app.loadFindChangeQuery( my_query_name, my_mode );    //쿼리를 세트
                if ( parseInt( app.version ) == 5 )
                    katakana2hex();     //카타카나에 매치하지 않는 버그 수정
                my_range_obj.changeGrep();
                // reDefineLog += '\r\n ' + my_query_name + ': ' + my_range_obj.changeGrep().length + ' 개';
            }
            else if ( my_mode == "glyph" )
            {
                my_mode = SearchModes.GLYPH_SEARCH;
                app.loadFindChangeQuery( my_query_name, my_mode );    //쿼리를 세트
                my_range_obj.changeGlyph();
                //reDefineLog += '\r\n ' + my_query_name + ': ' + my_range_obj.changeGlyph().length + ' 개';
            }
            else if ( my_mode == "object" )
            {
                my_mode = SearchModes.OBJECT_SEARCH;
                app.loadFindChangeQuery( my_query_name, my_mode );    //쿼리를 세트
                my_range_obj.changeObject();
                //reDefineLog += '\r\n ' + my_query_name + ': ' + my_range_obj.changeObject().length + ' 개', '스타일 재정의 결과';
            }
            else if ( my_mode == "transliterate" )
            {
                my_mode = SearchModes.TRANSLITERATE_SEARCH;
                app.loadFindChangeQuery( my_query_name, my_mode );    //쿼리를 세트
                my_range_obj.changeTransliterate();
                //reDefineLog += '\r\n ' + my_query_name + ': ' + my_range_obj.changeTransliterate().length + ' 개', '스타일 재정의 결과';
            }
            else MyErrorExit( "모드가 부정확합니다.：" + my_Queries[i][0] + "：" + my_query_name );
        } catch ( e )
        {
            reDefineLog = '에러가 발생했습니다.：' + my_Queries[i][0] + '：' + my_query_name + '\r' + e;
            errorCount++; totalError;
            return reDefineLog;
            //MyErrorExit("에러가 발생했습니다.：" + my_Queries[i][0] + "：" + my_query_name + "\r" + e);
        }
    }
    //alert(reDefineLog, docName.split('.')[0]+' 작업결과');
    //alert('작업이 완료되었습니다', '완료');
    if ( del_Image )
    {
        my_range_obj.characterStyles.itemByName( 'C_Image' ).baselineShift = 0;
        myDoc.characterStyles.itemByName( 'C_Image' ).remove();
    }
    reDefineLog = '작업이 완료되었습니다';
    return reDefineLog;
    //}/* for statement all documents */
}

function Make_Image()
{
    var my_Queries = [
        ["grep", "C_image", myNone],
        ["grep", "C_image_2", myNone],
        ["grep", "C_image_3", "MMI"],
        ["grep", "C_image_4", "C_NoBreak"],
        ["grep", "C_image_5", myNone]
    ];
    resultLog = 'C_Image 생성 결과'
    if ( RedefineCore( my_Queries, false ) < 0 )
        exit();
    if ( totalError )
        return resultLog;
    else
        return 'C_Image가 생성 되었습니다';
}

function Delete_Image()
{
    var my_Queries = [
        ["grep", "C_image_deselect", "C_Image"],
    ];
    resultLog = 'C_Image 삭제 결과'
    if ( RedefineCore( my_Queries, true ) < 0 )
        exit();
    if ( totalError )
        return resultLog;
    else
        return 'C_Image가 삭제되었습니다';
}

function Redefine_Style( my_Queries )
{
    resultLog = '스타일 재정의 결과'
    if ( RedefineCore( my_Queries, false ) < 0 )
        exit();
    if ( totalError )
        return resultLog;
    else
        return '성공적으로 완료되었습니다';
}

function RedefineCore( my_Queries, del_Image )
{
    var check = 0;
    var myDoc = null;
    var docNames = [];
    var log = '';
    try
    {
        myDoc = app.activeDocument;
        check = 1;
    } catch ( err ) { }
    if ( !check )
    {
        docNames = SelectTarget( true );
        if ( !docNames )
            return 0;
        else if ( docNames < 0 )
            return -1;
        else
            check = 2;
    }
    else
    {
        try { myDoc.filePath; }
        catch ( err )
        {
            alert( myDoc.name + '가 한 번도 저장되지 않았습니다.\r\n저장하신 후에 진행해주세요', '오류' );
            return -1;
        }
    }
    if ( check > 1 )
    {
        var res =
            "dialog { alignChildren: 'fill', \
                Success: Panel { orientation: 'row', alignChildren:'center', \
                    text: '문서들이 확인창 없이 자동저장 됩니다', \
                    s: StaticText { text: '진행하시겠습니까?' },\
                }, \
                buttons: Group { orientation: 'row', alignment: 'right', \
                    okBtn: Button { text:'OK', properties:{name:'ok'} }, \
                    cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                } \
            }";
        var result = new Window( res, '문서 자동저장 경고' );
        result.center();
        var select = result.show() - 1;
        if ( select )
            return -1;

        var bookDocuments = [];
        var bookDocumentsLength = docNames.length;
        var images, imageCount;
        for ( var i = 0 ; i < bookDocumentsLength ; i++ )
        {
            bookDocuments[i] = app.open( docNames[i], false );
            images = bookDocuments[i].allGraphics;
            var myMaster = bookDocuments[i].masterSpreads;
            var myMasterNumber = myMaster.count() - 1;
            var masterImage = 0;
            for ( ; myMasterNumber > -1 ; myMasterNumber-- )
            {
                imageNumber = myMaster[myMasterNumber].allGraphics;
                masterImage += imageNumber.length;
            }
            imageCount = images.length - masterImage;
            for ( var j = 0; j < imageCount ; j++ )
            {
                if ( images[j].itemLink.status != 1819109747 );
                else
                {
                    resultLog += '\r\n' + bookDocuments[i].name.split( '.' )[0] + '문서에 이미지링크 오류가 있습니다\r\n 링크를 재설정 해주세요';
                    totalError++;
                    bookDocuments[i].close( SaveOptions.YES, docNames[i] );
                    return resultLog;
                }
            }
            bookDocuments[i] = app.open( docNames[i], true );
            try
            {
                log = RedefineStyles( my_Queries, bookDocuments[i], del_Image );
                if ( totalError )
                    resultLog += '\r\n' + log;
            }
            catch ( err )
            {
                resultLog += '\r\n' + bookDocuments[i].name + ' : ' + '함수 오류';
                totalError++;
            };
            try { bookDocuments[i].close( SaveOptions.YES, docNames[i] ); }
            catch ( err )
            {
                resultLog += '\r\n' + bookDocuments[i].name + '실패';
                totalError++;
            }
        }
        app.activeBook.save();
        return resultLog;
    }
    else if ( check > 0 )
    {
        try { log = RedefineStyles( my_Queries, myDoc, del_Image ); }
        catch ( err )
        { log += '\r\n' + myDoc.name + ' : ' + '함수 오류'; }
        return log;
    }
}

function HighlightChecker( checker )
{
    var check = 0;
    var docNames = [];
    var myDoc = null;
    var log = null;
    var myTint = null;
    var myCell = null;
    
    try
    {
        myDoc = app.activeDocument;
        check = 1;
    } catch ( err ) { }
    if ( !check )
    {
        docNames = SelectTarget( true );
        if ( !docNames )
            return 0;
        else if ( docNames < 0 )
            return -1;
        else
            check = 2;
    }
    else
    {
        try { myDoc.filePath; }
        catch ( err )
        {
            alert( myDoc.name + '가 한 번도 저장되지 않았습니다.\r\n저장하신 후에 진행해주세요', '오류' );
            return -1;
        }
    }
    if ( check > 1 )
    {
        var res =
            "dialog { alignChildren: 'fill', \
                Success: Panel { orientation: 'row', alignChildren:'center', \
                    text: '문서들이 확인창 없이 자동저장 됩니다', \
                    s: StaticText { text: '진행하시겠습니까?' },\
                }, \
                buttons: Group { orientation: 'row', alignment: 'right', \
                    okBtn: Button { text:'OK', properties:{name:'ok'} }, \
                    cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                } \
            }";
        var result = new Window( res, '문서 자동저장 경고' );
        result.center();
        var select = result.show() - 1;
        if ( select )
            return -1;

        var bookDocuments = [];
        var bookDocumentsLength = docNames.length;
        for ( var i = 0 ; i < bookDocumentsLength ; i++ )
        {
            bookDocuments[i] = app.open( docNames[i], false );
            try { Checker( bookDocuments[i], checker ); }
            catch ( err )
            {
                resultLog += '\r\n' + bookDocuments[i].name + ' : ' + '함수 오류';
                totalError++;
            };
            try { bookDocuments[i].close( SaveOptions.YES, docNames[i] ); }
            catch ( err )
            {
                totalError++;
                resultLog += '\r\n' + bookDocuments[i].name + '실패';
                alert( docNames[i].name + '의 저장에 실패했습니다' );
            }
        }
        app.activeBook.save();
        if ( totalError )
            alert( resultLog, '오류' );
        else
            alert( '완료되었습니다', 'Checker' );
    }
    else if ( check > 0 )
    {

        try { Checker( myDoc, checker ); }
        catch ( err ) { alert( myDoc.name + ' : ' + '함수 오류', '오류' ); };
        if ( totalError ) alert( resultLog, '오류' );
        else alert( '완료되었습니다', 'Checker' );
    }
}

function Checker( myDoc, checker )
{
    //Target Objects
    var C_Image = myDoc.characterStyles.itemByName( 'C_Image' );
    var MMI = myDoc.characterStyles.itemByName( 'MMI' );
    var MMI_NoBold = myDoc.characterStyles.itemByName( 'MMI_NoBold' );
    var Iword = myDoc.characterStyles.itemByName( 'Iword' );
    var Iword_NoBold = myDoc.characterStyles.itemByName( 'Iword_NoBold' );
    var C_For_English = myDoc.characterStyles.itemByName( 'C_For_English' );
    var C_FontChange = myDoc.characterStyles.itemByName( 'C_FontChange' );
    var C_Singlestep = myDoc.characterStyles.itemByName( 'C_Singlestep' );
    var C_SingleStep = myDoc.characterStyles.itemByName( 'C_SingleStep' );
    var C_NoBreak = myDoc.characterStyles.itemByName( 'C_NoBreak' );

    var Cell_Table_H = myDoc.cellStyles.itemByName( 'Cell_Table_H' );
    var Cell_Table_H_K = myDoc.cellStyles.itemByName( 'Cell_Table_H_K' );
    var Cell_Number = myDoc.cellStyles.itemByName( 'Cell_Number' );
    var Cell_TableHead_Fill = myDoc.cellStyles.itemByName( 'Cell_TableHead-Fill' );
    var Cell_TableHead_Icon_Fill = myDoc.cellStyles.itemByName( 'Cell_TableHead-Icon-Fill' );
    var Cell_TableHead_Empty = myDoc.cellStyles.itemByName( 'Cell_TableHead-Empty' );

    //Swatches
    //var arrSwatches = myDocument.swatches;//All Swatches

    //Define MySwatches List
    var swatchList = [C_Image, Iword, Iword_NoBold, MMI, MMI_NoBold, C_For_English, C_Singlestep, C_SingleStep, C_NoBreak, Cell_Table_H, Cell_Table_H_K, Cell_Number, Cell_TableHead_Fill, Cell_TableHead_Icon_Fill, Cell_TableHead_Empty];//My Swatch List
    //var swatchList = [Cell_TableHead_Icon_Fill, Cell_TableHead_Empty];//My Swatch List
    if ( checker )
    {
        HighlightE( myDoc );
        Cell_None( myDoc.textFrames );
        for ( var i = 0; i < swatchList.length; i++ )
        {
            try
            {
                chkSwatchName( swatchList[i].name, myDoc );
                Highlight( swatchList[i], myDoc );
            }
            catch ( err ) { }
        }
    }
    else
    {
        UnHighlightE( myDoc );
        for ( var i = 0; i < swatchList.length; i++ )
        {
            UnHighlight( swatchList[i], myDoc );
            try
            {
                myDoc.colors.itemByName( swatchList[i].name ).remove();
                myFolder = new Folder( myDoc.filePath + '/Indesign_Highlightcheker/' + myDoc.name.split( '.' )[0] );
                myFolder.remove();
            }
            catch ( err ) { }
        }
    }
}

function Cell_None( myFrames )
{
    myFramesLength = myFrames.length;
    for ( var i = 0 ; i < myFramesLength ; i++ )
    {
        var myTables = myFrames[i].tables;
        var myTablesLength = myTables.length;
        for ( var j = 0 ; j < myTablesLength ; j++ )
        {
            var myCells = myTables[j].cells
            var myCellsLength = myCells.length;
            for ( var k = 0; k < myCellsLength ; k++ )
            {
                if ( myCells[k].appliedCellStyle.name == '[없음]' )
                    myCells[k].fillColor = null;
                else if ( myCells[k].appliedCellStyle.name == '[None]' )
                    myCells[k].fillColor = null;
            }
        }
    }
}

function HighlightE( myDoc )
{    //Do Highlight
    var myCell = myDoc.cellStyles;
    var cellLength = myDoc.cellStyles.length;
    for ( var i = 1; i < cellLength ; i++ )
    {
        try
        {
            if ( myCell[i].fillColor == null )
                myCell[i].fillColor = myDoc.colors.itemByName( 'Paper' );
        }
        catch ( err ) { }
    }
}

function UnHighlightE( myDoc )
{    //Do Highlight
    var myCell = myDoc.cellStyles;
    var cellLength = myDoc.cellStyles.length;
    for ( var i = 1; i < cellLength ; i++ )
    {
        try
        {
            if ( myCell[i].fillColor.name == 'Paper' )
                myCell[i].fillColor = null;
        }
        catch ( err ) { }
    }
}

function Highlight( target, myDoc )
{    //Do Highlight
    var istyle = ReplaceAll( target.name, '-', '_' );
    if ( istyle.search( 'Cell_' ) > -1 )
    {   //Cell스타일인 경우
        if ( typeof ( target.basedOn ) != 'string' )
            myDoc.xmlElements[0].xmlAttributes.add( istyle, 'inherit' );
        else if ( target.fillTint != 'NOTHING' && target.fillTint != 1851876449 && target.fillColor.name != 'None' )
            myDoc.xmlElements[0].xmlAttributes.add( istyle, target.fillTint.toString() );
        else
            myDoc.xmlElements[0].xmlAttributes.add( istyle, '0' );
        target.fillTint = 100;
        target.fillColor = myDoc.colors.itemByName( istyle );
    }
    else
    {    //Cell 스타일이 아닌 스타일
        try
        {
            target.underline = true;
            target.underlineColor = target.name;
            target.underlineWeight = THICK;
            target.underlineOffset = '-3pt';
            target.underlineTint = 100;
            if ( istyle == 'MMI' )
                myDoc.characterStyles.itemByName( 'C_Bold' ).underline = false
        } catch ( err ) { }
    }
}

function chkSwatchName( swchName, myDoc )
{  //Check Swatch Name
    swchName = ReplaceAll( swchName, '-', '_' );
    if ( myDoc.swatches.itemByName( swchName ) == null )
        createSwatch( swchName, eval( 'color' + swchName ), myDoc );
}

function createSwatch( swchName, swchColor, myDoc )
{    //Create Swatch
    var newSwatch = myDoc.colors.add();
    newSwatch.colorValue = swchColor;
    newSwatch.name = swchName;
}

function UnHighlight( target, myDoc )
{//Do Highlight
    try
    {
        var istyle = ReplaceAll( target.name, '-', '_' );
        if ( istyle.substr( 0, 5 ) == 'Cell_' )
        {//Cell스타일인 경우
            tint = myDoc.xmlElements[0].xmlAttributes.itemByName( istyle ).value;
            if ( tint == 'NOTHING' || tint == 1851876449 || tint == 0 )
            {
                target.fillTint = 0;
                target.fillColor = app.swatches.itemByName( 'None' );
            }
            else if ( tint == 'inherit' )
            {
                target.fillTint = NothingEnum.NOTHING;
                target.fillColor = NothingEnum.NOTHING;
            }
            else
            {
                target.fillTint = +tint;
                target.fillColor = myDoc.colors[0];
            }
            myDoc.xmlElements[0].xmlAttributes.itemByName( istyle ).remove();
            myDoc.colors.itemByName( istyle ).remove();
        }
        else
        {
            try
            {
                target.underlineWeight = '0.65pt';
                target.underlineOffset = '1.6pt';
                target.underline = false;
            } catch ( err ) { }
        }
    } catch ( err ) { }
}


function Paragraph_HighlightChecker(checker)
{
    var check = false;
    var docNames = [];
    var log = null;
    var myTint = null;
    var myCell = null;
    //var myFolder = new Folder( 'c:/Indesign_Highlightcheker/' );
    //myFolder.create();
    try {
        var myDoc = app.activeDocument;
        check = 1;
    } catch(err) {}
    if ( !check )
    {
        docNames = SelectTarget(1);
        if (!docNames)
            return 0;
        else if (docNames < 0)
            return -1;
        else
            check = 2;
    }
    else
    {
        try { myDoc.filePath; }
        catch(err) {
            alert(myDoc.name+'가 한 번도 저장되지 않았습니다.\n저장하신 후에 진행해주세요', '오류');
            return -1;
        }
    }

    if (check == 2)
    {
        var res =
            "dialog { alignChildren: 'fill', \
                Success: Panel { orientation: 'row', alignChildren:'center', \
                    text: '문서들이 확인창 없이 자동저장 됩니다', \
                    s: StaticText { text: '진행하시겠습니까?' },\
                }, \
                buttons: Group { orientation: 'row', alignment: 'right', \
                    okBtn: Button { text:'OK', properties:{name:'ok'} }, \
                    cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                } \
            }";
        var result = new Window (res, '문서 자동저장 경고');
        result.center();
        var select = result.show()-1;
        if (select) return -1;

        var bookDocuments = [];
        var bookDocumentsLength = docNames.length;
        for (var i=0 ; i<bookDocumentsLength ; i++)
        {
            bookDocuments[i] = app.open(docNames[i], false);
            try { Paragraph_Checker(bookDocuments[i], checker); }
            catch(err){ resultLog += '\n' + bookDocuments[i].name + ' : ' + '함수 오류'; totalError++; };
            try { bookDocuments[i].close(SaveOptions.YES, docNames[i]); }
            catch(err) { totalError++; resultLog += '\n' + bookDocuments[i].name +'실패'; alert(docNames[i].name + '의 저장에 실패했습니다' ); }
        }
        app.activeBook.save();
        if (totalError) alert(resultLog, '오류' );
        else alert('완료되었습니다', 'Checker');
    }
    else if (check == 1)
    {
        try { Paragraph_Checker(myDoc, checker); }
        catch(err) { alert( myDoc.name + ' : ' + '함수 오류', '오류'); };
        if (totalError)
            alert(resultLog, '오류' );
        else
            alert('완료되었습니다', 'Checker');
    }
}

function Paragraph_Checker(myDoc, checker)
{
    //Target Objects
    var swatchList = [];
    var paras = myDoc.paragraphStyles;
    var para_count = paras.length;
    for (var i  = 0 ; i < para_count ; i++)
    {
        if (paras[i].name.search('-App') > -1 || paras[i].name.search('-Recommend') > -1)
            swatchList.push(paras[i]);
    }
    var chars = myDoc.characterStyles;
    var chars_count = chars.length;
    for (var i = 0 ; i < chars_count ; i++)
    {
        if (chars[i].name.search('MMI') > -1 || chars[i].name.search('Iword') > -1)
            swatchList.push(chars[i]);
    }

    if(checker)
    {
        if (myDoc.swatches.itemByName('color_app') == null)
            createSwatch('color_app', color_app, myDoc);
        if (myDoc.swatches.itemByName('color_recommend') == null)
            createSwatch('color_recommend', color_recommend, myDoc);
        if (myDoc.swatches.itemByName('color_mmi') == null)
            createSwatch('color_mmi', color_mmi, myDoc);
        for (var i = 0; i < swatchList.length; i++)
        {
            try
            {   Paragraph_Highlight(swatchList[i], myDoc);    }
            catch (err) {}
        }
    }
    else
    {
        for (var i = 0; i < swatchList.length; i++) 
        {
            try
            {   Paragraph_UnHighlight(swatchList[i], myDoc);  }
            catch(err) {}
        }
        var color = myDoc.colors.itemByName('color_app');
        if (color != null)
            myDoc.colors.itemByName('color_app').remove();
        color = myDoc.colors.itemByName('color_recommend');
        if (color != null)
            myDoc.colors.itemByName('color_recommend').remove();
        color = myDoc.colors.itemByName('color_mmi');
        if (color != null)
            myDoc.colors.itemByName('color_mmi').remove();
    }
}

function Paragraph_Highlight(target, myDoc)
{
    var istyle = target.name;
    try
    {
        if (istyle.search('MMI') > -1 || istyle.search ('Iword') > -1)
        {
            target.fillColor = 'color_mmi';
            target.underline = NothingEnum.NOTHING;
            target.underlineColor = NothingEnum.NOTHING;
            target.underlineWeight = NothingEnum.NOTHING;
            target.underlineOffset = NothingEnum.NOTHING;
            target.underlineTint = NothingEnum.NOTHING;
            target.underlineType = NothingEnum.NOTHING;
        }
        else
        {
            target.underline = true;
            if (istyle.search('-App') > -1 )
                target.underlineColor = 'color_app';
            else if (istyle.search('-Recommend') > -1 )
                target.underlineColor = 'color_recommend';
            target.underlineWeight = '5pt';
            target.underlineOffset = '6pt';
            target.underlineTint = 100;
        }
    } catch (err) {}

}

function Paragraph_UnHighlight(target, myDoc) {//Do Highlight
    try{
        var istyle = target.name;
        if (istyle.search('MMI') > -1 || istyle.search ('Iword') > -1)
        {
            target.fillColor = NothingEnum.NOTHING;;// app.swatches[3];
        }
        else {
            try {
                target.underlineWeight = '0.65pt';
                target.underlineOffset = '1.6pt';
                target.underline = false;
            } catch(err) { }
        }
    }catch (err) {}
}


function Check_Changed_MMI( Checking )
{
    var myDoc = null;
    var check = 0;
    var docNames = [];
    try
    {
        myDoc = app.activeDocument;
        check = 1;
    } catch ( err ) { }
    if ( !check )
    {
        docNames = SelectTarget( true );
        if ( !docNames )
            return 0;
        else if ( docNames < 0 )
            return -1;
        else
            check = 2;
    }
    else
    {
        try { myDoc.filePath; }
        catch ( err )
        {
            alert( myDoc.name + '가 한 번도 저장되지 않았습니다.\r\n저장하신 후에 진행해주세요', '오류' );
            return -1;
        }
    }
    if ( check > 1 )
    {
        var res =
            "dialog { alignChildren: 'fill', \
                Success: Panel { orientation: 'row', alignChildren:'center', \
                    text: '문서들이 확인창 없이 자동저장 됩니다', \
                    s: StaticText { text: '진행하시겠습니까?' },\
                }, \
                buttons: Group { orientation: 'row', alignment: 'right', \
                    okBtn: Button { text:'OK', properties:{name:'ok'} }, \
                    cancelBtn: Button { text:'Cancel', properties:{name:'cancel'} } \
                } \
            }";
        var result = new Window( res, '문서 자동저장 경고' );
        result.center();
        var select = result.show() - 1;
        if ( select )
            return -1;

        var bookDocuments = [];
        var bookDocumentsLength = docNames.length;
        for ( var i = 0 ; i < bookDocumentsLength ; i++ )
        {
            bookDocuments[i] = app.open( docNames[i], false );
            try
            {
                if ( Checking )
                {
                    chkSwatchName( "ChangedMMI", bookDocuments[i] );
                    chkSwatchName( "ChangedMMI2", bookDocuments[i] );
                    Changed_MMI( bookDocuments[i], "MMI" );
                    Changed_MMI( bookDocuments[i], "MMI_NoBold" );
                }
                else
                {
                    UnChanged_MMI( bookDocuments[i], "MMI" );
                    UnChanged_MMI( bookDocuments[i], "MMI_NoBold" );
                    try { 
                    	bookDocuments[i].colors.itemByName( "ChangedMMI" ).remove();
                    	bookDocuments[i].colors.itemByName( "ChangedMMI2" ).remove(); 
                    } catch ( err ) { }
                }
            }
            catch ( err )
            {
                resultLog += '\r\n' + bookDocuments[i].name + ' : ' + '함수 오류';
                totalError++;
            }
            try { bookDocuments[i].close( SaveOptions.YES, docNames[i] ); }
            catch ( err )
            {
                totalError++;
                resultLog += '\r\n' + bookDocuments[i].name + '실패';
                alert( docNames[i].name + '의 저장에 실패했습니다' );
            }
        }
        app.activeBook.save();
        if ( totalError ) alert( resultLog, '오류' );
        else alert( '완료되었습니다', 'Checker' );
    }
    else if ( check > 0 )
    {
        try
        {
            if ( Checking )
            {
                chkSwatchName( "ChangedMMI", myDoc );
                chkSwatchName( "ChangedMMI2", myDoc );
                Changed_MMI( myDoc, "MMI" );
                Changed_MMI( myDoc, "MMI_NoBold" );
            }
            else
            {
                UnChanged_MMI( myDoc, "MMI" );
                UnChanged_MMI( myDoc, "MMI_NoBold" );
                try { 
                	myDoc.colors.itemByName( "ChangedMMI" ).remove();
                	myDoc.colors.itemByName( "ChangedMMI2" ).remove();
                } catch ( err ) { }
            }
        }
        catch ( err ) { alert( myDoc.name + ' : ' + '함수 오류', '오류' ); };
        if ( totalError )
            alert( resultLog, '오류' );
        else
            alert( '완료되었습니다', 'Checker' );
    }
}

function Changed_MMI( myDoc, myTag )
{
    app.findTextPreferences = NothingEnum.NOTHING;
    try
    {
        app.findTextPreferences.appliedCharacterStyle = myDoc.characterStyles.itemByName( myTag );
        var MMI_Value;
        var MMI_Changed = myDoc.characterStyles.itemByName( myTag + '_Changed' );
        var MMI_Changed2 = myDoc.characterStyles.itemByName( myTag + '_Changed2' );
        var detect = myDoc.findText();
        var detect_Length = detect.length;
        for ( var i = 0 ; i < detect_Length ; i++ ) {
            try {
			MMI_Value = detect[i].associatedXMLElements[0].xmlContent.contents;
			if ( MMI_Value.indexOf("$$$") != -1 ) {
            		detect[i].associatedXMLElements[0].xmlAttributes.itemByName( "CheckForBill" ).value;
            		detect[i].applyCharacterStyle( MMI_Changed2 );
            		//detect[i].associatedXMLElements[0].xmlContent.contents = MMI_Value.split("$$$").join("");
            	} else {
            		detect[i].associatedXMLElements[0].xmlAttributes.itemByName( "CheckForBill" ).value;
                	detect[i].applyCharacterStyle( MMI_Changed );
            	}
            } catch ( err ) { }
        }
        app.findTextPreferences = NothingEnum.NOTHING;
    }
    catch ( err ) { }
    return 1;
}

function UnChanged_MMI( myDoc, myTag )
{
    app.findTextPreferences = NothingEnum.NOTHING;
    try
    {
        app.findTextPreferences.appliedCharacterStyle = myDoc.characterStyles.itemByName( myTag + '_Changed' );
        var MMI_Origin = myDoc.characterStyles.itemByName( myTag );
        var detect = myDoc.findText();
        var detect_Length = detect.length;
        for ( var i = 0 ; i < detect_Length ; i++ )
            detect[i].applyCharacterStyle( MMI_Origin );
        app.findTextPreferences = NothingEnum.NOTHING;
        myDoc.characterStyles.itemByName( myTag + '_Changed' ).remove();
        
        app.findTextPreferences.appliedCharacterStyle = myDoc.characterStyles.itemByName( myTag + '_Changed2' );
        MMI_Origin = myDoc.characterStyles.itemByName( myTag );
        detect = myDoc.findText();
        detect_Length = detect.length;
        for ( var i = 0 ; i < detect_Length ; i++ )
            detect[i].applyCharacterStyle( MMI_Origin );
        app.findTextPreferences = NothingEnum.NOTHING;
        myDoc.characterStyles.itemByName( myTag + '_Changed2' ).remove();
    }
    catch ( err ) { }
    return 1;
}

function AddChangedStyle( myDoc )
{
    myDoc.characterStyles.add();
    var MMI_Changed = myDoc.characterStyles[myDoc.characterStyles.count() - 1];
    try {
        MMI_Changed.name = 'MMI_Changed';
        MMI_Changed.basedOn = myDoc.characterStyles.itemByName( 'MMI' );
        MMI_Changed.underline = true;
        MMI_Changed.underlineColor = "ChangedMMI";
        MMI_Changed.underlineWeight = ChangedTHICK;
        MMI_Changed.underlineOffset = '-4.5pt';
        MMI_Changed.underlineTint = 100;
    } catch ( err ) { MMI_Changed.remove(); }

	myDoc.characterStyles.add();
    MMI_Changed = myDoc.characterStyles[myDoc.characterStyles.count() - 1];
    try {
        MMI_Changed.name = 'MMI_Changed2';
        MMI_Changed.basedOn = myDoc.characterStyles.itemByName( 'MMI' );
        MMI_Changed.underline = true;
        MMI_Changed.underlineColor = "ChangedMMI2";
        MMI_Changed.underlineWeight = ChangedTHICK;
        MMI_Changed.underlineOffset = '-4.5pt';
        MMI_Changed.underlineTint = 100;
    } catch ( err ) { MMI_Changed.remove(); }
    	
    myDoc.characterStyles.add();
    var MMI_NoBold_Changed = myDoc.characterStyles[myDoc.characterStyles.count() - 1];
    try {
        MMI_NoBold_Changed.name = 'MMI_NoBold_Changed';
        MMI_NoBold_Changed.basedOn = myDoc.characterStyles.itemByName( 'MMI_NoBold' );
        MMI_NoBold_Changed.underline = true;
        MMI_NoBold_Changed.underlineColor = "ChangedMMI";
        MMI_NoBold_Changed.underlineWeight = ChangedTHICK;
        MMI_NoBold_Changed.underlineOffset = '-4.5pt';
        MMI_NoBold_Changed.underlineTint = 100;
    } catch ( err ) { MMI_NoBold_Changed.remove(); }
    	
	myDoc.characterStyles.add();
    MMI_NoBold_Changed = myDoc.characterStyles[myDoc.characterStyles.count() - 1];
    try {
        MMI_NoBold_Changed.name = 'MMI_NoBold_Changed2';
        MMI_NoBold_Changed.basedOn = myDoc.characterStyles.itemByName( 'MMI_NoBold' );
        MMI_NoBold_Changed.underline = true;
        MMI_NoBold_Changed.underlineColor = "ChangedMMI2";
        MMI_NoBold_Changed.underlineWeight = ChangedTHICK;
        MMI_NoBold_Changed.underlineOffset = '-4.5pt';
        MMI_NoBold_Changed.underlineTint = 100;
    } catch ( err ) { MMI_NoBold_Changed.remove(); }
}

// Log-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
// 실행할 항목 선택창 추가
// 태그, 언태그로 구분. radio버튼으로 변환
// 주석 처리
// UnTagging 부분의 remove 함수 추가
// 주석 수정, 오류 수정.
// 그룹화된 스타일을 위한 부분 추가
// 주석 수정, 오류 수정 - v3.3b
// remove 함수를 기본 함수로 변환 - v3.4a
// 추가할 스타일 선택창 오류 수정 - v3.5a
// 구조를 지우는 함수 추가 - v3.6
// OK, Cancel 버튼 정상 작동 -v3.8
// 선택 항목 텍스트 수정 -v3.8
// OK, Cancel 버튼 정상 작동 (v3.8에서는 cancel시 오작동)
// 선택창의 디폴트값 설정, 값을 설정하지 않아도 오류나지 않도록 수정 - v3.9
// 흐름 수정
// 그룹추가 함수 수정 v4.0
// CheckTag 함수 내에서 Compare 함수 호출 이전에 현재 태그 리스트를 갱신하도록 수정( 문자와 단락 스타일을 개별 추가 했을 경우 생기는 문제 수정) -v4.0_fix
// 구조 안의 MMI 태그를 카운팅하고 각각의 태그에 ID로 카운팅값을 추가하도록 수정 -v4.1a
// XML 가져오기 기능 추가 -v4.2a
// Compare 함수의 통합 (문자와 단락 각각의 함수를 합침) Add 함수 통합(역시 문자와 단락 각각의 함수를 통합) -v4.2b
// MappingToStyle() 함수 구현. 스타일에 태그를 매핑해준다. -v4.2c
// 검색 알고리즘 손볼 필요가 있음
// Add**Style 계열 함수들 수정 -v4.2d
// Find함수 추가 -v4.2e
// 스타일그룹만 있을경우 에러. 수정이 필요
// Find 함수에 태그카운터 값을 넘겨줄때 -1을 안해준것 주정 -v4.2f
// Find 함수에서 +1이 변수에 영향을 미치는 현상 수정
// MappingToStyle 함수에서 카운터가 잘못 되어 있던 것을 수정
// MappingToStyle 함수 알고리즘을 변경. nlogn 에서 n으로 -v.4.4
// RemakeStructure 함수 추가 -v4.5
// MappingToStyle 함수 수정 -v4.5
// RemakeStructure 함수 수정 -v4.5a
// 구조화 변수들 변경. 구조 수정  -v.4.6
// 구조화 변경. MappingToStyle 의 오류 수정( !Topic을 Topic으로 변경) -v4.6a
// ReTagging 할때 구조화의 오류 수정( 이미 태그가 있을시에 구조에서 오류나는 문제 수정 -v4.6b
// RemakeStructure 계열 함수들 수정 -v4.6e
// 오류 수정 -v4.7
// 초기 선택단계에 변화. 알고리즘 수정 -v4.7a
// Nomal 을 선택했을때 스타일에 매핑 안되는 문제 수정. -v4.7b
// PA_를 만들시에 Heading이 묶이는 것을 수정 -v4.7c
// PA 알고리즘 수정 -v 4.8
// PA 알고리즘수정, Export 기능 추가 -v4.9
// PA 알고리즘 수정(Empty가 포함되도록 수정) -v4.9a
// Restructure 계열 함수 상속관계로 구현, MMI 카운팅 삭제 -v5.0b
// PA 계열 알고리즘 수정 (Empty 관련 오류 수정, PA의 가운데 다른 스타일이 하나 있을경우 같이 묶게 변경) -v5.1a
// v5.1a 의 변경사항에 오류가 발생하여 롤백.
//  팬택 쪽의 스타일 추가 -v6.0
// 오류 수정 -v6.1b
// 함수들 간략화, 반복문 내의 메소드 변수로 지정.  함수 추가(rMoveElement, CheckHeading). Find 함수 제거. -v6.3
// Find, Compare 함수 복구( CS3에서 isvalid 가 적용되지 않음) -v7.0b
// XML내보내기 기능 수정
// 문자스타일 값을 지니는 텍스트의 공백 오류 체크 후 로그 작성 -v7.0f
// 흐름, 이름 정리 -v7.1a
// 스타일에 태그 매핑을 해주는 조건 수정 -v7.1a

// -v8,8 Image, Note, Icon을 잡아주는 if문의 오류 수정
// -v8.8 이미지의 문자스타일 지정에러일 경우 로그파일의 메세지 수정
// -v8.9 성공적으로 XML파일 추출시, Log파일 삭제
// -v9.0 MMI에 아이디 부여
// -v9.2 마스터페이지의 이미지는 크기조정에서 제외
// -v9.3 MakeCharErrorLog함수의 오류 수정. MMI_NoBold 에 특성을 하나 더 부여
// -v9.4 MakeCharErrorLog함수의 알고리즘 수정. 간소화
// -v9.4 MMI_NoBold 특성 삭제
// -v9.5 MakeCharErrorLog함수의 알고리즘 수정, 조건 강화, 출력문 수정
// -v9.6 MMI 만 별로도 XML추출기능 구현
// -v9.6 MMI 특성에 Chapter와 Heading 추가
// -v9.8 MMI특성에서 Heading 삭제, MMI 아이디 기준의 우선순위를 현재문서로 지정
// -v10.0 MMIFilter파일이 특정 폴더에 없는 경우 찾는 과정을 수행
// -v10.0 이미지를 가져올 때에 테두리값을 art로 고정
// -v10.1 import시 이미지 크기를 변환하지 않게 수정
// -v11.0 문법오류 수정. 필요없는 주석처리부분 삭제
// -v11.0 SelectTarget, ExportIDML 함수 추가
//
// -v11.2 RedefineStyles 함수 추가. 오류 수정
// -v11.3 함수 오류들 수정. cs5.5에서 참조 때문에 속도저하 문제 발견. book파일로 작업시 한번에 다 여는 방향으로 수정
// -v11.5 ►관련 선언 수정.
// -v 1.6 오류 검색을 해주는 기능을 따로 버튼으로 만듬. 북 파일이 다수일 경우 진행이 되지 않는 문제 해결

//14.2.11 import시 일치하지 않는 요소 삭제 옵션 취소.

//14.2.13 grep 'Refer_CharStyle_del8' 추가 - C_NoBreak가 쓰인 \r 제거
// empty 관련해서 임시 코드 작성

//14.3.1 iword 변환에 CHN_MMI 추가

// 함수들
function CallDate()
{
    var time = new Date;
    var cYear = MakeZero( time.getFullYear( cYear ) );
    var cMonth = MakeZero( time.getMonth() + 1 );
    var cDate = MakeZero( time.getDate() );
    var cHours = MakeZero( time.getHours() );
    var cMinutes = MakeZero( time.getMinutes() );
    var cSeconds = MakeZero( time.getSeconds() );

    return cYear + cMonth + cDate + cHours + cMinutes + cSeconds;
}
function MakeZero( timeData )
{
    timeData = ( timeData >> 1 | timeData >> 2 ) < 3 ? '0' + timeData.toString() : timeData.toString();
    return timeData;
}
function RandomNumber( min, max )
{
    return min + parseInt( Math.random() * ( max + 1 - min ) );
}

function SortNumber( a, b )
{
    return a - b;
}

function ReplaceAll( sValue, param1, param2 )
{
    return sValue.split( param1 ).join( param2 );
}


//그랩 세팅
function grep_setting(findwhat, appliedpara, appliedchar, changeto, changepara, changechar)
{
    try
    {
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
        return 1;
    }
    catch (ex)
    {
        alert('' + ex, 'grep_setting 함수 오류');
        return -1;
    }
}

//그랩 변경
function change_grep(doc, findwhat, appliedpara, appliedchar, changeto, changepara, changechar)
{
    try
    {
        if (grep_setting(findwhat, appliedpara, appliedchar, changeto, changepara, changechar) == -1)
            return -1;
        else
            doc.changeGrep();
        return 1;
    }
    catch (ex) 
    {
        alert('' + ex, 'change_grep 함수 오류');
        return -1;
    }
}

function Trim(stringToTrim, token) {
    stringToTrim = TrimEnd(TrimStart(stringToTrim, token), token);
    var reg = new RegExp('[\u0004\u0016\u0018\u0019\u000d\ufeff\ufffc]', 'g');
    stringToTrim = stringToTrim.replace(reg,"");
    
    return stringToTrim;
}

function TrimEnd(stringToTrim, token) {
    if (token == null || token == "\\s+") {
        token = "\\s+";
    }

    var reg = new RegExp(token + '$', 'g');
    return stringToTrim.replace(reg,"");
}

function TrimStart(stringToTrim, token) {
    if (token == null || token == "\\s+") {
        token = "\\s+";
    }

    var reg = new RegExp('^' + token, 'g');
    return stringToTrim.replace(reg,"");
}

function GetHashCode(str) {
    var hash = 0, c;
    if (str.length == 0) return hash;
    for (i = 0; i < str.length; i++) {
        c = str.charCodeAt(i) * (i + 1);
        hash += c; //((hash<<5)-hash)+c;
        //hash = hash & hash; // Convert to 32bit integer
    }
    return hash;
}

function TrimCharacterStyle(myDoc) {
    var allCharacterStyles = myDoc.allCharacterStyles;
    var none = allCharacterStyles[0];
    app.findTextPreferences = NothingEnum.NOTHING;
    for (var i = 1 ; i < allCharacterStyles.length ; i++) {
        app.findTextPreferences.appliedCharacterStyle = allCharacterStyles[i];
        var detectTexts = myDoc.findText();
        var detect_Length = detectTexts.length;
        if (detect_Length > 0) {
            for (var j = 0 ; j < detect_Length ; j++) {
                var text = detectTexts[j];
                var contents = text.contents;
                if (typeof(contents) == 'string') {
                    if (contents.length > 2) {
                        var contents2 = Trim(contents);
                        if (contents != contents2) {
                            var k = 0;
                            for (; k < text.contents.length ; k++) {
                                var c = text.characters[k];
                                if (c.contents == Trim(c.contents))
                                    break;
                            }
                            for (var l = 0 ; l < k ; l++) {
                                var c = text.characters[l];
                                c.appliedCharacterStyle = none;
                            }
                            k = text.contents.length - 1;
                            for (; k >= 0 ; k--) {
                                var c = text.characters[k];
                                if (c.contents == Trim(c.contents))
                                    break;
                            }
                            for (var l = k ; l < text.contents.length ; l++) {
                                var c = text.characters[l];
                                c.appliedCharacterStyle = none;
                            }
                        }
                    }
                }
            }
        }
    }
    app.findTextPreferences = NothingEnum.NOTHING;
}

function pi_to_breaks_and_specials(doc) {
    try {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
        var rootElement = doc.xmlElements.item(0);
        var targets = rootElement.evaluateXPathExpression("//processing-instruction('InDesign')");

        for (var i = 0; i < targets.length; i++) {
            var myPI = targets[i];
            switch (myPI.data) {
                case "5341706E" : myPI.storyOffset.contents = SpecialCharacters.autoPageNumber; break;
                case "534E706E" : myPI.storyOffset.contents = SpecialCharacters.nextPageNumber; break;
                case "5350706E" : myPI.storyOffset.contents = SpecialCharacters.previousPageNumber; break;
                case "53736E4D" : myPI.storyOffset.contents = SpecialCharacters.sectionMarker; break;
                case "53666E4D" : myPI.storyOffset.contents = SpecialCharacters.footnoteSymbol; break;
                case "53444870" : myPI.storyOffset.contents = SpecialCharacters.discretionaryHyphen; break;
                case "534E6268" : myPI.storyOffset.contents = SpecialCharacters.nonbreakingHyphen; break;
                case "53526974" : myPI.storyOffset.contents = SpecialCharacters.rightIndentTab; break;
                case "53496874" : myPI.storyOffset.contents = SpecialCharacters.indentHereTab; break;
                case "53425253" : myPI.storyOffset.contents = SpecialCharacters.endNestedStyle; break;
                case "00000009" : myPI.storyOffset.contents = "\t"; break;
                default: myPI.storyOffset.contents = parseInt("0x" + myPI.data);
            }
            myPI.remove();
        }
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    } catch (ex) {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
        throw ex;
    }
}

function specials_to_pi(doc) {
    try {
        var findWhatArray = [
            "~N", // autoPageNumber
            "~X", // nextPageNumber
            "~V", // previousPageNumber
            "~x", // sectionMarker
            "~F", // footnoteSymbol
            "~-", // discretionaryHyphen
            "~~", // nonbreakingHyphen
            "~y", // rightIndentTab
            "~i", // indentHereTab
            "~h", // endNestedStyle
            "\t"  // tab
        ];
        var changeToArray = [
            "5341706E", // autoPageNumber
            "534E706E", // nextPageNumber
            "5350706E", // previousPageNumber
            "53736E4D", // sectionMarker
            "53666E4D", // footnoteSymbol
            "53444870", // discretionaryHyphen
            "534E6268", // nonbreakingHyphen
            "53526974", // rightIndentTab
            "53496874", // indentHereTab
            "53425253", // endNestedStyle
            "00000009"  // tab
        ];
    
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.findChangeGrepOptions.includeFootnotes = false;
        app.findChangeGrepOptions.includeHiddenLayers = false;
        app.findChangeGrepOptions.includeLockedLayersForFind = false;
        app.findChangeGrepOptions.includeLockedStoriesForFind = false;
        app.findChangeGrepOptions.includeMasterPages = false;

        for (var i = 0; i < findWhatArray.length; i++) {
            app.findGrepPreferences.findWhat = findWhatArray[i];
            var matches = doc.findGrep();
            for ( j = 0; j < matches.length; j++) {
                var item = matches[j];
                item.select();
                var theElem = app.selection[0].associatedXMLElements[0];
                var thePI = theElem.xmlInstructions.add("InDesign", changeToArray[i]);
                thePI.move(LocationOptions.BEFORE, app.selection[0]);
                app.selection[0].remove();
            }
        }
        app.findGrepPreferences = app.changeGrepPreferences = null;
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    } catch (ex) {
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
        throw ex;
    }
}

function findchangeMMI ( myDoc ) {
    app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
    app.findTextPreferences.findWhat = "$$$";
    app.changeTextPreferences.changeTo = "";
    myDoc.changeText();
}

function updateCross ( myDoc ) {
    myDoc.updateCrossReferences();
}