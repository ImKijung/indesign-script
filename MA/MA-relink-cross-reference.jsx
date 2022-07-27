// 1. 상호참조를 찾는다.
// 2. 상호참조의 대상 문서와 앵커값을 얻는다.
// 3. 상호참조를 삭제한다.
// 4. 2번 값을 토대로 상호참조를 다시 연결한다.
// 4-1. 목적지가 외부인지 내부인지 확인한다. parent 문서와 현재 문서의 이름 비교

var docs = app.documents;
$.writeln("Start");
// 연결할 파일이 모두 열려있는지 확인
var count = 0;
for (var i=0;i<docs.length;i++){
	// alert(docs[i].name);
	var openDocs = app.documents.everyItem().getElements();
	app.activeDocument = openDocs[openDocs.length-1];
	var doc = app.activeDocument;
	if (doc.name == "003_Basics.indd") {
		count ++;
	}
	if (doc.name == "004_Applications.indd") {
		count ++;
	}
	if (doc.name == "005_Settings.indd") {
		count ++;
	}
	if (doc.name == "006_Appendix.indd") {
		count ++;
	}
	if (doc.name == "003_Getting started.indd") {
		count ++;
	}
	if (doc.name == "006_Usage notices.indd") {
		count ++;
	}
	if (doc.name == "007_Appendix.indd") {
		count ++;
	}
	else continue;
}

if (!(count == 4 || count == 5)) {
	alert ("모든 문서가 열리지 않았습니다. 다시 확인하고 실행하세요.");
	exit();
}

if (app.books.length == 0) {
	alert ("북 파일이 열려 있지 않습니다. 다시 확인하고 실행하세요.");
	exit();
}
// 경고메시지 비활성화
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

// 79431 모든 상호 참조 업데이트
app.scriptMenuActions.itemByID(79431).invoke();

for (var i=0;i<docs.length;i++){
	var openDocs = app.documents.everyItem().getElements();
	app.activeDocument = openDocs[openDocs.length-1];
	//alert(openDocs[openDocs.length-1].name);
	relinkCrossreference();
}
alert ("모든 상호 참조를 재연결했습니다.");

// 경고메시지 활성화
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

// #target estoolkit
// //do stuff
// app.quit();

// relinkCrossreference();
function relinkCrossreference() {
	var doc = app.activeDocument;
	var myCross = doc.crossReferenceSources.everyItem().getElements();
	var myRef = doc.hyperlinks;
	for (var i=0; i<myRef.length;i++) {
		// app.select(myRef[i].sourceText);
		// selection = app.selection[0];
		try {
			var dest = myRef[i].destination;
			if (dest instanceof HyperlinkTextDestination) {
				app.select(myRef[i].source.sourceText);
				var mySourceText = app.selection[0];
				var targetDocument = myRef[i].destination.parent.name;
				var myDestiny = myRef[i].destination.name;
				$.writeln(i+1 + " " + mySourceText.contents + " - " + targetDocument + " - " + myDestiny);
				myRef[i].remove();
				mySourceText.remove();

				var xRefFormat = doc.crossReferenceFormats.itemByName("전체 단락"); 
				var mySource = doc.crossReferenceSources.add(mySourceText, xRefFormat);
				// var myHypDest = doc.hyperlinkTextDestinations.itemByName("Navigation bar (soft buttons)");
				// var mySource = doc.crossReferenceSources.item(i);
				var target = app.documents.itemByName(targetDocument);
				var myHypDest = target.hyperlinkTextDestinations.itemByName(myDestiny);
				// var highlight = HyperlinkAppearanceHighlight.INVERT;
				doc.hyperlinks.add(
					{ source: mySource, 
					  destination : myHypDest,
					  highlight : HyperlinkAppearanceHighlight.INVERT
					}
				);
				//document.hyperlinks.add ( { source : document.hyperlinkPageItemSources.add ( { sourcePageItem : selectedItem } ) , destination : document.hyperlinkURLDestinations.add (  urlDestString ) , borderStyle : HyperlinkAppearanceStyle.SOLID , hidden : false , highlight : HyperlinkAppearanceHighlight.INVERT } );
			}
		} catch(e) {
			// alert(e);
			$.writeln(e);
		}
	}
	doc.save();
}