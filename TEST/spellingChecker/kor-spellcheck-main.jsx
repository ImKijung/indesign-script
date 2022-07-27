#targetengine "session";
#include 'functions/restix.jsx'
#include 'functions/json2.jsx'

var docs = app.documents;
if (docs.length == 0) {
	alert("인디자인 문서가 열려있지 않습니다. 문서를 하나라도 열어 놓은 상태에서 실행하세요.");
	exit();
}

var w = new Window ("palette", "AST | 한글 맞춤법 검사기 ver.0.0.1");
panel = w.add ('panel {text: ""}');
panel.orientation = "row";
var btn_01 = panel.add("button", [0, 0, 110, 25], "검사 시작");
var btn_02 = panel.add("button", [0, 0, 110, 25], "이전 문장");
var btn_03 = panel.add("button", [0, 0, 110, 25], "선택 문장 검사");
panel01 = w.add ('panel {text: "원문"}');
var myText01 = panel01.add('edittext', [0, 0, 351, 150], '', {multiline:true});
	myText01.text = "인디자인 스크립트 : 한글 맞춤법 검사\r네이버 맞춤법 검사기를 이용합니다. 맞춤범 검사 내용이 100% 맞다고 확신할 수 없습니다. 매뉴얼의 작업 환경에 따라서 적용 여부는 작업자 본인이 판단하시기 바랍니다."
panel02 = w.add ('panel {text: "제안"}');
var myText02 = panel02.add('edittext', [0, 0, 351, 150], '', {multiline:true});
	myText02.text = "제안 내용으로 출력된 내용은 네이버 맞춤법 검사기를 통해 나온 결과물입니다. 교정결과에 오류가 있을 수 있습니다."
var count = 0;

w.show();

btn_01.onClick = function() {
	count = count + 1;
	var doc = app.activeDocument;
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
	app.findChangeGrepOptions.properties = {
		includeFootnotes:false,
		includeHiddenLayers:false,
		includeLockedLayersForFind:false,
		includeLockedStoriesForFind:false,
		includeMasterPages:false,
	}

	app.findGrepPreferences.findWhat = "[^\\r]+";

	var myFound = doc.findGrep();
	
	
	for (i=0; i<myFound.length; i++) {
		var mySel = myFound[count-1].contents;
		mySel = mySel.replace("\n", "");
		// $.writeln(mySel);
		panel.text = myFound.length + " 문장 중 " + count + " 번째 문장 검사";
		spellingCheckWhole(myFound, count, mySel, myText01, myText02);
		count++;
		if (count == myFound.length) {
			count = 0; //카운트 초기화
			panel.text = myFound.length + " 문장 중 " + count+1 + " 번째 문장 검사";
		}
	}
	
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
}

btn_02.onClick = function() {
	var doc = app.activeDocument;
	count = count - 1;
	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
	app.findChangeGrepOptions.properties = {
		includeFootnotes:false,
		includeHiddenLayers:false,
		includeLockedLayersForFind:false,
		includeLockedStoriesForFind:false,
		includeMasterPages:false,
	}

	app.findGrepPreferences.findWhat = "[^\\r]+";

	var myFound00 = doc.findGrep();
	
	for (j=count; j<myFound00.length; j--) {
		var mySel = myFound00[count-1].contents;
		mySel = mySel.replace(/\n/g, "");
		// $.writeln(mySel);
		// $.writeln(count, mySel);
		panel.text = myFound00.length + " 문장 중 " + count + " 번째 문장 검사";
		spellingCheckWhole(myFound00, count, mySel, myText01, myText02);
		count --;
	}
	// $.writeln(count);

	app.findGrepPreferences = NothingEnum.nothing;
	app.changeGrepPreferences = NothingEnum.nothing;
}

btn_03.onClick = function() {
	var mySel = app.selection[0].contents;
	mySel = mySel.replace(/\n/g, "");
	spellingCheck(mySel, myText01, myText02);
}

function spellingCheckWhole(myFound, count, mySel, myText01, myText02) {
	var myparam = encodeURI(mySel);
	var request = {
		url:"https://m.search.naver.com/p/csearch/ocontent/util/SpellerProxy?color_blindness=0&q=" + myparam
	}

	var response = restix.fetch(request);
	//$.writeln(response.toSource());

	if (response.error) {
		myText01.text = "응답 오류: " + response.error + "\r오류 메시지: " + response.errorMsg
	}
	//$.writeln("Response HTTP Status: " + response.httpStatus);
	//$.writeln("Response Body: " + response.body);

	var newJson = JSON.parse(response.body);
	//$.writeln(newJson.toSource());
	var original = newJson.message.result.origin_html.replace(/\<span class\=\'result_underline\'\>/g, "/*").replace(/\<\/span\>/g, "*/");
	var modify = newJson.message.result.html.replace(/\<em class\=\'green_text\'\>/g, "/*").replace(/\<em class\=\'red_text\'\>/g, "/*").replace(/\<em class\=\'blue_text\'\>/g, "/*").replace(/\<\/em\>/g, "*/");
	try {
		if (newJson.message.result.errata_count != 0) {
			myFound[count-1].select();
			//app.activeWindow.activePage = myFound[count-1].parentTextFrames[0].parentPage;
			app.activeDocument.layoutWindows[0].zoomPercentage = 300;
			myText01.text = original;
			myText02.text = modify;
			exit();
		} else {
			myText01.text = original;
			myText02.text = "제안 사항 없음";
		}
	} catch(e) {
		alert(e.line + ":" + e);
	}
}

function spellingCheck(mySel, myText01, myText02) {
	var myparam = encodeURI(mySel);
	var request = {
		url:"https://m.search.naver.com/p/csearch/ocontent/util/SpellerProxy?color_blindness=0&q=" + myparam
	}

	var response = restix.fetch(request);
	//$.writeln(response.toSource());

	if (response.error) {
		myText01.text = "응답 오류: " + response.error + "\r오류 메시지: " + response.errorMsg
	}
	//$.writeln("Response HTTP Status: " + response.httpStatus);
	//$.writeln("Response Body: " + response.body);

	var newJson = JSON.parse(response.body);
	//$.writeln(newJson.toSource());
	var original = newJson.message.result.origin_html.replace(/\<span class\=\'result_underline\'\>/g, "/*").replace(/\<\/span\>/g, "*/");
	var modify = newJson.message.result.html.replace(/\<em class\=\'green_text\'\>/g, "/*").replace(/\<em class\=\'red_text\'\>/g, "/*").replace(/\<em class\=\'blue_text\'\>/g, "/*").replace(/\<\/em\>/g, "*/");
	try {
		if (newJson.message.result.errata_count != 0) {
			myText01.text = original;
			myText02.text = modify;
			exit();
		} else {
			myText01.text = original;
			myText02.text = "제안 사항 없음";
		}
	} catch(e) {
		alert(e.line + ":" + e);
	}
}