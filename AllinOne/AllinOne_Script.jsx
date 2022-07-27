#targetengine "session";
#include "Manual/all_save_close_SERIAL2.jsx"
#include "Manual/Batch_Export_IDML.jsx"
#include "Manual/copy2Links.jsx"
#include "Manual/Batch_fit-image_to_frame.jsx"
#include "Manual/Auto-Metadata.jsx"
#include "Manual/Update_alert_links.jsx"
#include "Manual/Hyperlink_apply_underline.jsx"
#include "Manual/insert_unicode.jsx"
#include "Manual/change_language_parastyle.jsx"
#include "Manual/toc_contents_export.jsx"
#include "Manual/index_export.jsx"
#include "Manual/unused_style_delete.jsx"
#include "Manual/change_font.jsx"
#include "Manual/trynow_contents_export.jsx"
#include "Manual/style_mod.jsx"
#include "Manual/highlight_no_trans.jsx"
#include "Manual/Rename_Links_Script.jsx"
#include "Manual/RtoL_Style_Changer.jsx"

AllinOne_Script();

function AllinOne_Script() {
    var w = new Window ("palette","All in One Script v2.0.2");  
		//w.orientation = "row";
	var group1 = w.add("group")
		group1.add("statictext", undefined, "적용기준 - [f]:폴더단위 | [x]:옵션선택 | [1]:1개의 문서 | [b]:북파일단위 | [+]:열려있는 모든 문서")
	var group2 = w.add("group")
		group2.orientation = "row";
	var panel_1 = group2.add ("panel");
	var panel_2 = group2.add ("panel");
	var panel_3 = group2.add ("panel");
    
    btn_02 = panel_1.add ("button", [0,0,200,30], "인디자인 Export IDML[f]");
	btn_04 = panel_1.add ("button", [0,0,200,30], "GREP 매니저[x]");
	btn_25 = panel_1.add ("button", [0,0,200,30], "미주구주 영단어 변경[x]");
	btn_07 = panel_1.add ("button", [0,0,200,30], "이미지 경로 설정[x]");
	btn_03 = panel_1.add ("button", [0,0,200,30], "이미지 모아주기[x]");
	btn_28 = panel_1.add ("button", [0,0,200,30], "이미지 파일명 변환[1]");
	btn_10 = panel_1.add ("button", [0,0,200,30], "*수정된 이미지 업데이트[+]");
    btn_06 = panel_1.add ("button", [0,0,200,30], "프레임을 이미지(100%)에 맞춤[1]");
	btn_05 = panel_1.add ("button", [0,0,200,30], "이미지를 프레임에 맞춤[1]");

    btn_08 = panel_2.add ("button", [0,0,200,30], "*001chapter_메타정보입력[1]");
	btn_21 = panel_2.add ("button", [0,0,200,30], "*C_image 메타정보 입력-2017[b]");
    btn_09 = panel_2.add ("button", [0,0,200,30], "C_image 메타정보 입력-2016[1]");
	btn_24 = panel_2.add ("button", [0,0,200,30], "Try Now URL 리스트 만들기[f]");
    btn_11 = panel_2.add ("button", [0,0,200,30], "*Reference to Underline[+]");
	btn_14 = panel_2.add ("button", [0,0,200,30], "RtoL 유니코드 입력[1]");
    btn_15 = panel_2.add ("button", [0,0,200,30], "LtoR 유니코드 입력[1]");
	btn_01 = panel_2.add ("button", [0,0,200,30], "SERIAL2 저장 닫기[+]");
	btn_12 = panel_2.add ("button", [0,0,200,30], "열린 문서 저장 닫기[+]");
    
    btn_17 = panel_3.add ("button", [0,0,200,30], "단락/문자 스타일 목록[1]");
	btn_19 = panel_3.add ("button", [0,0,200,30], "단락스타일 콘텐츠 추출[1]");
	btn_20 = panel_3.add ("button", [0,0,200,30], "문자스타일 콘텐츠 추출[x]");
	btn_23 = panel_3.add ("button", [0,0,200,30], "폰트 변경하기[1]");
    btn_16 = panel_3.add ("button", [0,0,200,30], "단락스타일 언어 변경[1]");
    btn_26 = panel_3.add ("button", [0,0,200,30], "단락/문자스타일 기본 변환기[1]");
	btn_29 = panel_3.add ("button", [0,0,200,30], "RtoL 단락/문자스타일 변환기[1]");
	btn_27 = panel_3.add ("button", [0,0,200,30], "No_trans 색상 적용[x]");
	btn_22 = panel_3.add ("button", [0,0,200,30], "사용하지 않는 스타일 삭제하기[x]");

	// var mainScriptpath = "C:/Program Files (x86)/Adobe/Adobe InDesign CS5.5/Scripts/Scripts Panel/Manual/"
	var scriptPath = app.activeScript.path;
	var mainScriptpath = scriptPath + "/Manual/";
	// $.writeln(mainScriptpath);

//----------------------------------------------------------------------------------------------------
	btn_01.onClick = function () {
		all_save_close_SERIAL2();
	}
//----------------------------------------------------------------------------------------------------
	btn_02.onClick = function () {
		Batch_Export_IDML();
	}
//----------------------------------------------------------------------------------------------------
	btn_03.onClick = function () {
		w.close();
		copy2Links(allinOnePath);
	}
//----------------------------------------------------------------------------------------------------
	btn_04.onClick = function () {
		w.close();
		var myScriptPath = mainScriptpath + "grep_query_manager.jsx";
		var myScriptFile = new File (myScriptPath);
		$.evalFile(myScriptFile);
	}
//----------------------------------------------------------------------------------------------------
	btn_05.onClick = function () {
		Batch_fit_image_to_frame();
	}
//----------------------------------------------------------------------------------------------------
	btn_06.onClick = function () {
		Batch_fit_image_size100_frame_to_content();
	}
//----------------------------------------------------------------------------------------------------
	btn_07.onClick = function () {
		w.close();
		var myScriptPath = mainScriptpath + "Update_path_links.jsx";
		var myScriptFile = new File (myScriptPath);
		$.evalFile(myScriptFile);
	}
//----------------------------------------------------------------------------------------------------
	btn_08.onClick = function () {
		AutoMetadata();
	}
//----------------------------------------------------------------------------------------------------
	btn_09.onClick = function () {
		var myScriptPath = mainScriptpath + "ApplyALTfromXMP.jsx";
		var myScriptFile = new File (myScriptPath);
		$.evalFile(myScriptFile);
	}
//----------------------------------------------------------------------------------------------------
	btn_10.onClick = function () {
		Update_alert_links();
	}
//----------------------------------------------------------------------------------------------------
	btn_11.onClick = function () {
		Hyperlink_apply_underline();
	}
//----------------------------------------------------------------------------------------------------
	btn_12.onClick = function () {
		var docs = app.documents;
		for (var i = docs.length-1; i >= 0; i--) {
			docs[i].close(SaveOptions.YES);
		}
		alert("모든 문서를 저장했습니다.");
	}
//----------------------------------------------------------------------------------------------------

	btn_14.onClick = function () {
		insertRtoL();
	}
//----------------------------------------------------------------------------------------------------
	btn_15.onClick = function () {
		insertLtoR();
	}
//----------------------------------------------------------------------------------------------------
	btn_16.onClick = function () {
		change_language_parastyle();
	}
//----------------------------------------------------------------------------------------------------
	btn_17.onClick = function () {
		w.close();
		var myScriptPath = mainScriptpath + "List_ps_cs.jsx";
		var myScriptFile = new File (myScriptPath);
		$.evalFile(myScriptFile);
	}
//----------------------------------------------------------------------------------------------------
	btn_19.onClick = function () {
		toc_contents_export();
	}
//----------------------------------------------------------------------------------------------------
	btn_20.onClick = function () {
		index_export();
	}
//----------------------------------------------------------------------------------------------------
	btn_21.onClick = function () {
		var myScriptPath = mainScriptpath + "Alt_text_2017/ApplyALTfromXMP_book.jsx";
		var myScriptFile = new File (myScriptPath);
		$.evalFile(myScriptFile);
	}
//----------------------------------------------------------------------------------------------------
	btn_22.onClick = function () {
		unused_style_delete();
	}
//----------------------------------------------------------------------------------------------------
	btn_23.onClick = function () {
		ChagneFont();
	}
//----------------------------------------------------------------------------------------------------
	btn_24.onClick = function () {
		w.close();
		trynow_contents_export();
	}
//----------------------------------------------------------------------------------------------------
	btn_25.onClick = function () {
		w.close();
		var myScriptPath = mainScriptpath + "swipe_eng_word.jsx";
		var myScriptFile = new File (myScriptPath);
		$.evalFile(myScriptFile);
	}
//----------------------------------------------------------------------------------------------------
	btn_26.onClick = function () {
		w.close();
		styleMod();
	}
//----------------------------------------------------------------------------------------------------
	btn_27.onClick = function () {
		w.close();
		highlight_no_trans_main();
	}
//----------------------------------------------------------------------------------------------------
	btn_28.onClick = function () {
		renameLinksScript();
	}
//----------------------------------------------------------------------------------------------------
	btn_29.onClick = function () {
		w.close();
		RtoL_Style_Changer()
	}
//----------------------------------------------------------------------------------------------------

	w.show();
}