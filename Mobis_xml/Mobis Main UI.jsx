#include "Reference/common.jsx";
#include "Functions/ExportFloatId.jsx";
#include "Functions/MakeXML.jsx";
#include "Functions/CheckAndMakeTags.jsx";
#include "Functions/CheckAndFixErrors.jsx";
#include "Functions/MakeFlatStructure.jsx";
#include "Functions/RemoveStructureAndTags.jsx";
#include "Functions/ControlImages.jsx";
#include "Functions/EditTOC.jsx";
#include "Functions/ExportIDML.jsx";
#include "Functions/EditHyperlink.jsx";
#include "Functions/EditIndex.jsx";
#include "Functions/ControlHighlight.jsx";
#include "Functions/ControlLayer.jsx";
#include "Functions/D-CaptionLabel.jsx";
#include "Functions/CopyNotes.jsx";
#include "Functions/otherOptions.jsx";

var errorMsgs = new Array();
var hasError = false;
var isStopped = false;
var CloseWithoutMsg = 'CloseWithoutMsg';

var imageSize = 100;
var folderPath = null;
var folderPath2 = null;
var linked = [];
var layerName = null;
var layerNames = [];
var layerVisibles = [];

var askSave = true;
var askBook = true;

var currentFolder = new File($.fileName).parent;
var referenceFolder = null;
if (currentFolder.name == 'Reference')
    referenceFolder = currentFolder;
else
    referenceFolder = new Folder(currentFolder.fsName + '\\Reference');
var indentXML = new File(referenceFolder.fsName + '\\indent_XML.xsl');
var inddFile = new File(referenceFolder.fsName + '\\ImageCharacterStyle.indd');
var DBxmlFile = new File(referenceFolder.fsName + '\\DB.xml');
DBxmlFile.open('r');
var DBxml = new XML(DBxmlFile.read());
DBxmlFile.close();
var imageStyleName = DBxml.xpath("/Root/CharacterStyles/*[@type='CharImage'][1]").attribute('name').toString();
var appImageStyleName = DBxml.xpath("/Root/CharacterStyles/*[@type='CharImage'][1]").attribute('globalName').toString();
var floatIdStyles = DBxml.xpath("/Root/ParagraphStyles/*[@type='floatId']");
var TOCremoveStyles = DBxml.xpath("/Root/CharacterStyles/*[@type='TOC'][@action='remove']");
var TOCNumberStyles = DBxml.xpath("/Root/CharacterStyles/*[@type='TOC'][@action='setNumber']");
//imageStyleName = 'C_Image'
var higthlightStyles = DBxml.xpath("//*[@type='highlight']");
var styledImageTables = DBxml.xpath("/Root/TableStyles/*[@type='useImageCharacterStyle']");
var checkImageStyled = DBxml.xpath("/Root/Configs/Config[@type='image'][@name='checkImageStyled'][@value='true'][1]")
if (checkImageStyled != null)
    checkImageStyled = checkImageStyled.attribute('value').toString() == 'true' ? true : false;

var mmiStyles = DBxml.xpath("/Root/CharacterStyles/*[@type='MMI']");
var mmiCount = 0;

var THICK = '5mm';//밑줄 두께
var ChangedTHICK = '5mm';//밑줄 두께
var SHIFT = '3pt';//베이스라인 높이

var progresspanel = new Window("window", "실행중");
with (progresspanel) { progresspanel.progressbar = add("progressbar", [12, 12, 400, 24], 0, 100); }
progressTitle = "실행 중...";


function SetGlobalVariableInitialize() {
    errorMsgs = new Array();
    hasError = false;
    isStopped = false;
    imageSize = 100;
    folderPath = null;
    folderPath2 = null;
    linked = [];
    layerName = null;
    layerNames = [];
    layerVisibles = [];
    askSave = true;
    askBook = true;
    mmiCount = 0;
    progressTitle = "실행 중...";
    progresspanel.text = progressTitle;
}


main_ui();

function main_ui() {
    var res =
        "dialog { alignChildren: 'fill', \
            OptionXML: Panel { alignChildren:'left', \
                text: 'XML', \
                GroupXML: Group { orientation: 'row', alignment: 'center', \
                    functionExportXML: Button { text: 'XML 추출', minimumSize: [205, 25] },\
                    functionImportXML: Button { text: 'XML 적용', minimumSize: [205, 25] }\
                } \
                GroupTagAndStructure: Group { orientation: 'row', alignChildren:'center', \
                    functionMakeFlatStructure: Button { text: 'XML 구조 및 태그 생성', minimumSize: [205, 25] },\
                    functionRemoveStructureAndTags: Button { text: 'XML 구조 및 태그 제거', minimumSize: [205, 25] }\
                } \
                GroupExport: Group { orientation: 'row', alignChildren:'center', \
                    functionExportIDML: Button { text: 'IDML 파일 내보내기', minimumSize: [205, 25] },\
                    functionExportFloatId: Button { text: 'Float ID 추출', minimumSize: [205, 25] }\
                } \
            },\
            OptionImage: Panel { alignChildren:'left', \
                text: '이미지', \
                GroupIamge: Group { orientation: 'row', alignChildren:'center', \
                    functionResizelImages: Button { text: '전체 이미지 사이즈 조정', minimumSize: [205, 25] },\
                    functionChangeImageFolder: Button { text: '전체 이미지 링크 재 설정', minimumSize: [205, 25] }\
                } \
                GroupIamge2: Group { orientation: 'row', alignChildren:'center', \
                    functionAddImageStyle: Button { text: '이미지 문자 스타일 추가 및 적용', minimumSize: [205, 25] },\
                    functionDelelteImageStyle: Button { text: '이미지 문자 스타일 삭제', minimumSize: [205, 25] }\
                } \
                GroupIamge3: Group { orientation: 'row', alignChildren:'center', \
                    functionCheckImageSize: Button { text: '이미지 사이즈 확인', minimumSize: [205, 25] },\
                    functionSetImageStyle: Button { text: '이미지 문자 스타일 편집', minimumSize: [205, 25] }\
                } \
            },\
            OptionHighlight: Panel { alignChildren:'left', \
                text: '하이라이트', \
                GroupHighlight: Group { orientation: 'row', alignChildren:'center', \
                    functionSetHighlight: Button { text: '하이라이트 적용', minimumSize: [205, 25] },\
                    functionResetHighlight: Button { text: '하이라이트 해제', minimumSize: [205, 25] }\
                } \
            },\
            OptionTOC: Panel { alignChildren:'left', \
                text: 'TOC', \
                GroupTOC: Group { orientation: 'row', alignChildren:'center', \
                    functionUpdateTOC: Button { text: '목차 업데이트', minimumSize: [205, 25] },\
                    functionUpdateTOCRtL: Button { text: '목차 업데이트(RtoL)', minimumSize: [205, 25] }\
                } \
                GroupTOC2: Group { orientation: 'row', alignChildren:'center', \
                    functionRemoveCharcterStyledText: Button { text: 'TOC에서 스타일된 텍스트 삭제', minimumSize: [205, 25] },\
                    functionDeleteIndexes: Button { text: '모든 색인 정보 삭제', minimumSize: [205, 25] }\
                } \
            }\
            OptionLayer: Panel { alignChildren:'left', \
                text: '레이어', \
                GroupLayer: Group { orientation: 'row', alignChildren:'left', \
                    functionHideLayer: Button { text: '레이어 숨기기', minimumSize: [205, 25] },\
                    functionShowLayer: Button { text: '레이어 보이기', minimumSize: [205, 25] }\
                } \
                GroupLayer2: Group { orientation: 'row', alignChildren:'left', \
                    functionAddLayer: Button { text: '레이어 추가', minimumSize: [205, 25] },\
                    functionDeleteLayer: Button { text: '레이어 삭제', minimumSize: [205, 25] }\
                } \
            }\
            Option: Panel { alignChildren:'left', \
                text: '그 외', \
                Option0: Group { orientation: 'row', alignChildren:'left', \
                    functionCaptionLabel: Button { text: '도안명 자동 표기', minimumSize: [205, 25] },\
                    userInteractionLevel: Button { text: '알람 및 경고창 설정 되돌리기', minimumSize: [205, 25] }\
                } \
                Option1: Group { orientation: 'row', alignChildren:'left', \
                    functionleftToright: Button { text: '바인딩/스토리/테이블 방향 전환', minimumSize: [205, 25] },\
                    allDocusaveClose: Button { text: '열린 문서 저장 닫기', minimumSize: [205, 25] }\
                } \
                Option2: Group { orientation: 'row', alignChildren:'left', \
                    insertUnicode: Button { text: '유니코드 입력기', minimumSize: [205, 25] },\
                } \
            }\
        }";
    var handler = null;
    var selectWindow = new Window(res, '기능 선택 창');
    selectWindow.OptionXML.GroupXML.functionExportXML.onClick = function () {
        handler= functionExportXML;
        selectWindow.close(1);
    };
    selectWindow.OptionXML.GroupXML.functionImportXML.onClick = function () {
        handler= functionImportXML;
        selectWindow.close(1);
    };
    selectWindow.OptionXML.GroupTagAndStructure.functionMakeFlatStructure.onClick = function () {
        handler= functionMakeFlatStructure;
        selectWindow.close(1);
    };
    selectWindow.OptionXML.GroupTagAndStructure.functionRemoveStructureAndTags.onClick = function () {
        handler= functionRemoveStructureAndTags;
        selectWindow.close(1);
    };
    selectWindow.OptionImage.GroupIamge.functionResizelImages.onClick = function () {
        handler= functionResizelImages;
        selectWindow.close(1);
    };
    selectWindow.OptionImage.GroupIamge.functionChangeImageFolder.onClick = function () {
        handler= functionChangeImageFolder;
        selectWindow.close(1);
    };
    selectWindow.OptionImage.GroupIamge2.functionAddImageStyle.onClick = function () {
        handler= functionAddImageStyle;
        selectWindow.close(1);
    };
    selectWindow.OptionImage.GroupIamge2.functionDelelteImageStyle.onClick = function () {
        handler= functionDelelteImageStyle;
        selectWindow.close(1);
    };
    selectWindow.OptionImage.GroupIamge3.functionCheckImageSize.onClick = function () {
        handler= functionCheckImageSize;
        selectWindow.close(1);
    };
    selectWindow.OptionImage.GroupIamge3.functionSetImageStyle.onClick = function () {
        handler= functionSetImageStyle;
        selectWindow.close(1);
    };
    selectWindow.OptionTOC.GroupTOC.functionUpdateTOC.onClick = function () {
        handler= functionUpdateTOC;
        selectWindow.close(1);
    };
    selectWindow.OptionTOC.GroupTOC.functionUpdateTOCRtL.onClick = function () {
        handler= functionUpdateTOCRtL;
        selectWindow.close(1);
    };
    selectWindow.OptionTOC.GroupTOC2.functionRemoveCharcterStyledText.onClick = function () {
        handler= functionRemoveCharcterStyledText;
        selectWindow.close(1);
    };
    selectWindow.OptionTOC.GroupTOC2.functionDeleteIndexes.onClick = function () {
        handler= functionDeleteIndexes;
        selectWindow.close(1);
    };
    selectWindow.OptionHighlight.GroupHighlight.functionSetHighlight.onClick = function () {
        handler= functionSetHighlight;
        selectWindow.close(1);
    };
    selectWindow.OptionHighlight.GroupHighlight.functionResetHighlight.onClick = function () {
        handler= functionResetHighlight;
        selectWindow.close(1);
    };
    selectWindow.OptionXML.GroupExport.functionExportIDML.onClick = function () {
        handler= functionExportIDML;
        selectWindow.close(1);
    };
    selectWindow.OptionXML.GroupExport.functionExportFloatId.onClick = function () {
        handler= functionExportFloatId;
        selectWindow.close(1);
    };
    selectWindow.OptionLayer.GroupLayer.functionHideLayer.onClick = function () {
        handler= functionHideLayer;
        selectWindow.close(1);
    };
    selectWindow.OptionLayer.GroupLayer.functionShowLayer.onClick = function () {
        handler= functionShowLayer;
        selectWindow.close(1);
    };
    selectWindow.OptionLayer.GroupLayer2.functionAddLayer.onClick = function () {
        handler= functionAddLayer;
        selectWindow.close(1);
    };
    selectWindow.OptionLayer.GroupLayer2.functionDeleteLayer.onClick = function () {
        handler= functionDeleteLayer;
        selectWindow.close(1);
    };
    selectWindow.Option.Option0.functionCaptionLabel.onClick = function () {
        handler = functionCaptionLabel;
        selectWindow.close(1);
    };
    selectWindow.Option.Option0.userInteractionLevel.onClick = function () {
        handler = show_alert_dialog;
        selectWindow.close(1);
    };
    selectWindow.Option.Option1.functionleftToright.onClick = function () {
        handler = functionleftToright;
        selectWindow.close(1);
    };
    selectWindow.Option.Option1.allDocusaveClose.onClick = function () {
        handler = allDocusaveClose;
        selectWindow.close(1);
    };
    selectWindow.Option.Option2.insertUnicode.onClick = function () {
        // handler = insertUnicode;
        selectWindow.close(3);
    };

    selectWindow.helpTip = '필요한 기능을 클릭하면 다음 단계로 넘어갑니다';
    //selectWindow.Option.Option1.function_errorcheck.helpTip = '에러 체크 관련 기능';
    //selectWindow.Option.Option2.function_image.helpTip = '이미지 관련 기능';
    selectWindow.center();
    
    try {
        var select = selectWindow.show();
        if (select == 1) {
            SetGlobalVariableInitialize();
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
            var result = handler();
            
            if (errorMsgs.length > 0) {
                var title = '메세지';
                var useErrorIcon = false;
                if (hasError) {
                    title = '오류 메세지';
                    useErrorIcon = true;
                }
                var msg = null;
                if (hasError && errorMsgs.length > 10) {
                    msg = '다수의 오류가 발생하였습니다. 로그 파일을 참고해주십시오.';
                } else {
                    msg = errorMsgs.join('\r\n');
                }
                alert(msg, title, useErrorIcon);
                
                var time = CallDate();
                var logFolder = new Folder('Mobis_XML_Log');
                var log = new File(logFolder.fsName + '\\Extend Script ErrorLog ' + time + '.txt');
                if (!logFolder.exist) {
                    if (!logFolder.create())
                        log = new File('Extend Script ErrorLog ' + time + '.txt');
                }
                
                log.open('w');
                log.write(errorMsgs.join('\r\n'));
                log.close();
                if (hasError)
                    log.execute();
            } else {
                if (hasError) {
                    alert('프로그램 실행 도중 알수 없는 오류가 발생하였습니다.', '오류 메세지 공란', true);
                } else if (result == CloseWithoutMsg) {
                    //just close.
                } else if (isStopped) {
                    alert('프로그램이 중지되었습니다.', '중지 메세지');
                } else if (result) {
                    alert('완료하였습니다', '완료');
                } else {
                    alert('Unkwon Error. Return value is false.', 'Error', true);
                }
            }
            
            if (isStopped) {
                isStopped = false;
                return main_ui();
            }

        } else if (select == 3) {
            InsertUnicdoefunc();
            
        } else if (select == 0)
            return main_ui();
    } catch (ex) { 
         alert('' + ex, '오류', true);
    }
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}

function alertDelay() {
    alert('구현 중...', '미구현 메세지');
    isStopped = true;
    return CloseWithoutMsg;
}

function show_alert_dialog () {
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
    alert('모든 알람 및 경고창이 보이도록 설정합니다.', '완료');
    isStopped = true;
    return CloseWithoutMsg;
}
