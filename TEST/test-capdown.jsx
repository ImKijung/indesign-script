var doc = app.documents;
var doc = app.activeDocuments;
for (var j=0; j< doc.length; j++) {
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
	app.findGrepPreferences.findWhat = "(\\:) ([A-Z])|(\\:) (Į)|(\\:) ([\\x{0400}-\\x{04F0}])";
	var myFound = doc.findGrep();
	var counter = 0;

	// for (var i=0; i < myFound.length; i++) {
	// 	myFound[i].changecase(ChangecaseMode.LOWERCASE);
	// 	counter ++;
	// } alert (doc[j].name + "문서에서 " + counter + " 개의 문장을 확인했습니다.")
	app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
} 