// 북 파일 오픈 확인
if (app.books.length == 0) {
	alert("북 파일을 열려있지 않습니다.");
	exit();
}
// 경고메시지 비활성화
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var myBook = app.activeBook; // 현재 열려있는 북 파일
var myBookContents = myBook.bookContents.everyItem().getElements();
var master;

for(var i=0; i<myBookContents.length; i++) {
	var myPath = myBookContents[i].fullName;
	var myFileName = myBookContents[i].name;
	var myFile = File(myPath);
	// $.writeln(myFileName);
	if (myFileName.search('Cover') > 0 || myFileName.search('TOC') > 0) { //커버, 목차 파일 제외
		continue;
	} else {
		var doc = app.open(myFile);
		if (i == 2 || i == 10 || i == 18) { // 파일 순서에 따라 마스터 페이지 적용
			master = "D-Index";
			// $.writeln(master + ": 1111111");
			applyMasters(doc, master);
		} else if (i == 3 || i == 11 || i == 19) {
			master = "E-Index";
			// $.writeln(master + ": 2222222");
			applyMasters(doc, master);
		} else if (i == 4 || i == 12 || i == 20) {
			master = "F-Index";
			// $.writeln(master + ": 3333333");
			applyMasters(doc, master);
		} else if (i == 5 || i == 13 || i == 21) {
			master = "G-Index";
			// $.writeln(master + ": 4444444");
			applyMasters(doc, master);
		} else if (i == 6 || i == 14 || i == 22) {
			master = "H-Index";
			// $.writeln(master + ": 5555555");
			applyMasters(doc, master);
		} else if (i == 7 || i == 15 || i == 23) {
			master = "I-Index";
			// $.writeln(master + ": 6666666");
			applyMasters(doc, master);
		} else if (i == 8 || i == 16 || i == 24) {
			master = "J-Index";
			// $.writeln(master + ": 7777777");
			applyMasters(doc, master);
		} else if (i == 9 || i == 17 || i == 25) {
			master = "K-Index";
			// $.writeln(master + ": 8888888");
			applyMasters(doc, master);
		}
		doc.close(SaveOptions.YES); // 문서 저장
	}
}
// 경고메시지 활성화
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
alert("와아아아안료");

function applyMasters(doc, master) { //마스터 페이지 적용 코드
	// $.writeln(master);
	var myPages = doc.pages;
	for (var j=0; j<myPages.length; j++) {
		myPages[j].appliedMaster = doc.masterSpreads.item(master);
	}
}