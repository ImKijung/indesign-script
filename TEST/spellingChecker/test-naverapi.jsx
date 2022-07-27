#include 'functions/restix.jsx'
#include 'functions/json2.jsx'

var myContents = "아버지가가방에 들어가시다.";
var myparam = encodeURI(myContents);
var request = {
	url:"https://openapi.naver.com/v1/search/errata.json?query=" + myparam,
	method: "GET",
	headers: [
		{name: "Content-Type", value: "application/x-www-form-urlencoded;charset=UTF-8"},
		{name: "X-Naver-Client-Id", value: "HFKRxYDMhoqc6uWCQE7O"},
		{name: "X-Naver-Client-Secret", value: "kV9XLS9MWD"},
	]
}

var response = restix.fetch(request);
//$.writeln(response.toSource());

if (response.error) {
	myText01.text = "응답 오류: " + response.error + "\r오류 메시지: " + response.errorMsg
}
//$.writeln("Response HTTP Status: " + response.httpStatus);
//$.writeln("Response Body: " + response.body);

var newJson = JSON.parse(response.body);
$.writeln(newJson.toSource());