#include 'functions/restix.jsx'
#include 'functions/json2.jsx'

// var param = JSON.stringify({
// 		language: 'en',
// 		contents: 'An inertial force is a force that resists a change in velocity of an object.',
// 	});

// var language = "en";
// var contents = "An inertial force is a force that resists a change in velocity of an object.";
// var param = "language=" + language + "&contents=" + contents;

var request = {
	url: "https://vadim-berman-tisane-v1.p.rapidapi.com/parse",
	method: "POST",
	body: JSON.stringify({
		text: "An inertial force is a force that resists a change in velocity of an object.",
		topics: ["physics"],
	}),
	headers: [
		{name: "content-type", value: "application/json"},
		{name: "x-rapidapi-key", value: "e9d2566806msh6055ca83ceb34a3p126915jsn20e78916ac66"},
		{name: "x-rapidapi-host", value: "vadim-berman-tisane-v1.p.rapidapi.com"}
	],
}


var response = restix.fetch(request);
if (response.error) {
	$.writeln("Response Error: " + response.error);
	$.writeln("Response errorMsg: " + response.errorMsg);
}
$.writeln("Response HTTP Status: " + response.httpStatus);
$.writeln("Response Body: " + response.body);