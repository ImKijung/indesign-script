#include 'functions/restix.jsx'
#include 'functions/json2.jsx'

// var myparam = app.selection[0].contents;
// $.writeln(myparam);
var lang = "en-GB";
var texts = "To initialize all system sound settings, press Reset.";
var param = "?text=" + texts + "&language=" + lang;
var request = {
	url:"https://grammarbot.p.rapidapi.com/check" + encodeURI(param),
	// command: "posts",
	method: "POST",
	headers: [
		{name: "Content-Type", value: "application/x-www-form-urlencoded"},
		{name: "x-rapidapi-key", value: "e9d2566806msh6055ca83ceb34a3p126915jsn20e78916ac66"},
		{name: "x-rapidapi-host", value: "grammarbot.p.rapidapi.com"}
	]
}

var response = restix.fetch(request);

if (response.error) {
	$.writeln("Response Error: " + response.error);
	$.writeln("Response Error Message: " + response.errorMsg);
}
//$.writeln("Response HTTP Status: " + response.httpStatus);
//$.writeln("Response Body: " + response.body);

var newJson = JSON.parse(response.body);
$.writeln(newJson.toSource());


// Example Request
//  var request = {
//  	url:"String",
//  	command:"String", // defaults to ""
//  	port:443, // defaults to ""
//  	method:"GET|POST", // defaults to GET
//  	headers:[{name:"String", value:"String"}], // defaults to []
//  	body:"" // defaults to ""
// }

// var testCase = "// Creating a resource";
// request = {
// 	url:"https://jsonplaceholder.typicode.com",
// 	command:"posts", 
// 	method:"POST",
	
// 	body: JSON.stringify({
//       title: 'foo',
//       body: 'bar',
//       userId: 1
//     }),

//     headers: [{name:"Content-type", value:"application/json; charset=UTF-8"}]
// }
// var response = restix.fetch(request);
// logResponse (response, testCase) ;

// function logResponse (response, testCase) {
// 	$.writeln(testCase + "  ----");
// 	if (response.error) {
// 		$.writeln("Response Error: " + response.error);
// 		$.writeln("Response errorMsg: " + response.errorMsg);
// 	}
// 	$.writeln("Response HTTP Status: " + response.httpStatus);
// 	$.writeln("Response Body: " + response.body);
// }