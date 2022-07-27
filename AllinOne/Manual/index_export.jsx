//Index Export Main
#targetengine "session";
#include "index_contents_docu_export.jsx"
#include "index_contents_book_export.jsx"

function index_export() {
	var w = new Window ("palette","문자 스타일 콘텐츠 목록 추출");
	//var panel_1 = w.add ("panel");
	btn_01 = w.add ("button", [0,0,190,30], "문서 단위 추출");
	btn_02 = w.add ("button", [0,0,190,30], "북 단위 추출");


	btn_01.onClick = function () {
		index_contents_docu_export();
	}
	btn_02.onClick = function () {
		index_contents_book_export();
	}
	w.show();
}