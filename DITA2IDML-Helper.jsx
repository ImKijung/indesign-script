app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var files;
// var folder = Folder("C:/ProteanUI/idmlModule/out"); 
if (app.documents.length > 0) {
	for (var i=app.documents.length -1; i >= 0; i--) {
		app.documents[i].close(SaveOptions.YES);
	}
}
var folder = Folder.selectDialog( "IDML 문서가 있는 폴더를 선택하세요." );
if (folder != null) {
    files = GetFiles(folder);
    if (files.length > 0) {
        for (var u = 0; u < files.length; u++) { 
            var doc = app.open(files[u], true);
            // files[i].remove();
            var inddName = getFileName(files[u].fullName).replace("idml", "indd");
            myfunc(doc);
            doc.save(new File(folder + "/" + inddName));
            // doc.close();
        }
    }
    else { 
        alert("Found no match"); 
    }
} else {
	exit();
}

makeBook();
updateCrossAll();

alert("완료");
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;

function makeBook() {
	var folder = app.activeDocument.filePath;
	var bookFileName = prompt ("", "", "북 파일명을 입력하세요.");
	var docs = app.documents;
	if (bookFileName != null) {
		var myBookFileName = folder + "/"+ bookFileName + ".indb";
		myBookFile = new File( myBookFileName );
		myBook = app.books.add( myBookFile );
		myBook.automaticPagination = false;
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("001_Cover") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("002_TOC") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("003_Basics") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("003_Getting_started") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("004_Applications") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("005_Settings") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("006_Appendix") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("006_Usage_notices") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("007_Copyright") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("007_Appendix") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		for (i=0; i<docs.length; i++) {
			if (docs[i].name.indexOf("008_Copyright") != -1) {
				myFile = File(folder + "/" + docs[i].name);
				myBook.bookContents.add( myFile );
			}
		}
		myBook.automaticPagination = true;
		
		for (var j=0; j<docs.length; j++) {
			if (docs[j].name.indexOf("007_Copyright") != -1 || docs[j].name.indexOf("008_Copyright") != -1) {
				//007 Copyright 텍스트 프레임 하단 정렬, 002 목차 업데이트
				app.activeDocument = docs[j];
				var myDoc = app.activeDocument;
				var myPage = myDoc.pages.item(-1);
				myPage.appliedMaster = myDoc.masterSpreads.item("D-Copyright");
				
				var myTF = myPage.textFrames[0];
				myTF.textFramePreferences.verticalJustification = VerticalJustification.BOTTOM_ALIGN;
				myTF.geometricBounds = [25.0000000018403,15.9999983344792,276.999999999461,193.999998347582];
				
				var myBook = app.activeBook;
				var bookPath = myBook.filePath;
				var lang = bookPath.toString().match(/[^\/\/]+$/g);
				_backCover(lang);
				myDoc.save();
			} 
		}
		for (var j=0; j<docs.length; j++) {
			if (docs[j].name.indexOf("002_TOC") != -1) {
				// 목차 업데잍, 목차 제목을 페이지에 개체로 표시하기
				app.activeDocument = docs[j];
				var myDoc = app.activeDocument;
				myDoc.textPreferences.smartTextReflow = false;
				myDoc.pages.item(0).textFrames[0].select();
				try{  
					app.scriptMenuActions.itemByID(71442).invoke();
				} catch(e) { 
					alert("목차 업데이트 에러!!");
				}
				var mstItem = myDoc.pages[0].appliedMaster.pageItems;
				mstItem[0].override(myDoc.pages[0]);
				// 스타일 재정의
				var Pages = myDoc.pages;
				for (var y=0; y<Pages.length; y++) {
					var tf = Pages[y].textFrames;
					for (z=0; z<tf.length; z++) {
						var para = tf[z].paragraphs;
						for (w=0; w<para.length; w++) {
							para[w].clearOverrides(OverrideType.ALL);
							para[w].appliedCharacterStyle = app.activeDocument.characterStyles[0];
						}
						for (w=0; w<para.length; w++) {
							para[w].clearOverrides(OverrideType.ALL);
							para[w].appliedCharacterStyle = app.activeDocument.characterStyles[0];
						}
						for (w=0; w<para.length; w++) {
							para[w].clearOverrides(OverrideType.ALL);
							para[w].appliedCharacterStyle = app.activeDocument.characterStyles[0];
						}
					}
				}
				var paras = myDoc.stories.everyItem().paragraphs.everyItem().getElements();
				for (var w=0; w<paras.length; w++) {
					paras[w].clearOverrides(OverrideType.ALL);
					paras[w].clearOverrides(OverrideType.ALL);
					paras[w].clearOverrides(OverrideType.ALL);
				}
				myDoc.save();
			}
			var delGrepStyle = docs[j].paragraphStyles.itemByName("Navigation")
			var nestGrep = delGrepStyle.nestedGrepStyles;
			if (nestGrep.length > 0) {
				for (var g=nestGrep.length - 1; g>=0;g--) {
					nestGrep[g].remove();
				}
			}
			docs[j].save();
		}
		myBook.save();
	} else {
		alert("북 파일명을 입력하지 않았습니다. 모든 과정을 취소합니다.");
		exit();
	}
}


function getFileName(value) {
	var path = value.toString();
	var lastIndex = path.lastIndexOf("/");
	return path.slice(lastIndex + 1);
}

function GetFiles(theFolder) { 
    var files = [], 
    fileList = theFolder.getFiles(), i, file; 

    for (i = 0; i < fileList.length; i++) { 
        file = fileList[i]; 
        if (file instanceof Folder) { 
        } 
        else if (file instanceof File && file.name.match(/\.idml$/i)) {
            files.push(file); 
        } 
    } 
    return files;
}

function myfunc (myDoc) { 
	var page;
	myDoc.textPreferences.deleteEmptyPages = true;
	myDoc.recompose();
	
	var imageCount = GetImageCountWithoutMaster(myDoc);

	var myGraphics = myDoc.allGraphics;
	for (var j = 0; j < imageCount; j++) {
		var myGraphic = myGraphics[j];
		myGraphic.parent.fit(FitOptions.FRAME_TO_CONTENT);
	}
	applyMaster(myDoc);
}

function makeTOC(myDoc) {
	var tocStyles = myDoc.tocStyles;
	var nTocStyles = tocStyles.length;

	for (var h = 1; h < nTocStyles; h++) {
		myDoc.createTOC(tocStyles[h], true);
	}
}

function GetImageCountWithoutMaster(myDoc) {
    try {
        var myMaster = myDoc.masterSpreads;
        var myMasterNumber = myMaster.count();
        var masterImages = 0;
        for (var v = 0 ; v < myMasterNumber ; v++) {
            imageNumber = myMaster[v].allGraphics;
            masterImages += imageNumber.length;
        }

        return myDoc.allGraphics.length - masterImages;
    } catch (ex) {
        alert(ex);
    }
}

function applyMaster(myDoc) {
	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
    app.findTextPreferences.findWhat = "";
	app.findTextPreferences.appliedParagraphStyle = "Chapter";
	var find = myDoc.findText();
	for (var r = 0; r < find.length; r++) {
		var myPage = find[r].parentTextFrames[0].parentPage;
		myPage.appliedMaster = myDoc.masterSpreads.item("B-Chapter");
	}
    app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;

	//app.findTextPreferences.findWhat = "";
	//app.findTextPreferences.appliedParagraphStyle = "Heading1-H3";
	//var find2 = doc.findText();
	//var cPage = find2[0].parentTextFrames[0];
	//cPage.textFramePreferences.verticalJustification = VerticalJustification.bottomAlign;
	//cPage.parentPage.appliedMaster = doc.masterSpreads.item("D-Copyright");

	app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
}
function updateCrossAll() {
	// 1. 상호참조를 찾는다.
	// 2. 상호참조의 대상 문서와 앵커값을 얻는다.
	// 3. 상호참조를 삭제한다.
	// 4. 2번 값을 토대로 상호참조를 다시 연결한다.
	// 4-1. 목적지가 외부인지 내부인지 확인한다. parent 문서와 현재 문서의 이름 비교

	//연결할 파일이 모두 열려있는지 확인
	var docs = app.documents;
	for (var y=0;y<docs.length;y++){
		relinkCrossreference(docs[y]);
	}
	// alert ("모든 상호 참조를 재연결했습니다.");

	// relinkCrossreference();
	function relinkCrossreference(doc) {
		// var doc = app.activeDocument;
		var myCross = doc.crossReferenceSources.everyItem().getElements();
		var myRef = doc.hyperlinks;
		for (var c=0; c<myRef.length;c++) {
			try {
				var dest = myRef[c].destination;
				if (dest instanceof HyperlinkTextDestination) {
					// app.select(myRef[i].source.sourceText);
					// var mySourceText = app.selection[0];
					var mySourceText = myRef[c].source.sourceText;
					var targetDocument = myRef[c].destination.parent.name;
					var myDestiny = myRef[c].destination.name;
					// $.writeln(i+1 + " " + mySourceText.contents + " - " + targetDocument + " - " + myDestiny);
					myRef[c].remove();
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
				}
				else continue;
			} catch(e) {
				alert(e);
			}
		}
		doc.save();
	}
}

function _backCover(lang) {
	try {
		if (lang == "tr-TR") {
			var Conts = app.activeDocument.pages[0].textFrames[0].paragraphs;
			var count = 0;
			for (var j=0;j<Conts.length;j++) {
				if (Conts[j].appliedParagraphStyle.name == "Heading1-H3") {
					break
				}
				count ++;
			}
			// 1st 25,16,36,194.749999999968
			// 2nd 42.2499999997305,16,64.9999999997305,194.749999999968
			// 3rd 70.3999999997305,16,93.1499999997305,194.749999999968
			var turTf01 = app.activeDocument.textFrames.add({
				geometricBounds : [25.0000000018403,15.9999983344793,36,193.999998347582],
			});
			turTf01.textFramePreferences.firstBaselineOffset = FirstBaseline.ASCENT_OFFSET;
			var turTf02 = app.activeDocument.textFrames.add({
				geometricBounds : [42.2499999997305,15.9999983344793,64.9999999997305,193.999998347582],
			});
			turTf02.textFramePreferences.firstBaselineOffset = FirstBaseline.ASCENT_OFFSET;
			var turTf03 = app.activeDocument.textFrames.add({
				geometricBounds : [70.3999999997305,15.9999983344793,93.1499999997305,193.999998347582],
			});
			turTf03.textFramePreferences.firstBaselineOffset = FirstBaseline.ASCENT_OFFSET;
			Conts = app.activeDocument.pages[0].textFrames[3].paragraphs;
			// $.writeln(count);
			for (var k=count-1; k>=0; k--) {
				if (k == 0) {
					Conts[k].remove();
				} else if (k == 4) {
					Conts[k].select();
					app.cut();
					turTf03.insertionPoints[0].select();
					app.paste();
				} else if (k == 3) {
					Conts[k].select();
					app.cut();
					turTf02.insertionPoints[0].select();
					app.paste();
				} else if (k <= 2) {
					Conts[k].select();
					app.cut();
					turTf01.insertionPoints[0].select();
					app.paste();
				}
			}
		} else if (lang == "da-DK" || lang == "nl-NL" || lang == "fi-FI" || lang == "it-IT" || lang == "nn-NO" || lang == "pl-PL" || lang == "es-ES" || lang == "sv-SE" || lang == "zh-TW") {
			var Conts = app.activeDocument.pages[0].textFrames[0].paragraphs;
			var count = 0;
			for (var j=0;j<Conts.length;j++) {
				if (Conts[j].appliedParagraphStyle.name == "Heading1-H3") {
					break
				}
				count ++;
			}
			// $.writeln(count);
			var copyTf = app.activeDocument.textFrames.add({
				geometricBounds : [25.0000000018403,15.9999983344792,145.166666666397,193.999999999936],
			});
			copyTf.textFramePreferences.firstBaselineOffset = FirstBaseline.ASCENT_OFFSET;
			var foundSharp;
			for (var k=count-1; k>=0; k--) {
				Conts = app.activeDocument.pages[0].textFrames[1].paragraphs;
				if (Conts[k].appliedParagraphStyle.name == "Heading1") {
					Conts[k].remove();
				} else {
					Conts[k].select();
					app.cut();
					copyTf.insertionPoints[0].select();
					app.paste();
				}
			}
		} else if (lang == "ko-KR") {
			var curDoc = app.activeDocument;
			
			app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
			
			app.findGrepPreferences.findWhat = "^NO-TITLE\\r";
			app.findGrepPreferences.appliedParagraphStyle = "Heading1";
			var find2 = curDoc.findGrep();
			for (var i=0;i<find2.length;i++) {
				find2[i].remove();
			}
			app.findGrepPreferences = app.changeGrepPreferences = NothingEnum.NOTHING;
			
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;
			
			app.findTextPreferences.findWhat = "";
			app.findTextPreferences.appliedParagraphStyle = "Heading1-H3";
			app.changeTextPreferences.appliedParagraphStyle = "Description-NoHTML";
			curDoc.changeText();

			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;

			app.findTextPreferences.findWhat = "";
			app.findTextPreferences.appliedParagraphStyle = "Heading3";
			app.changeTextPreferences.appliedParagraphStyle = "Description-NoHTML";
			curDoc.changeText();
			
			app.findTextPreferences = app.changeTextPreferences = NothingEnum.NOTHING;

			if (curDoc.pages.length > 1) {
				curDoc.pages[0].textFrames[0].textFramePreferences.insetSpacing = [ 4, 4, 4, 4 ];
				curDoc.pages[0].textFrames[0].strokeWeight = 0.5;
				curDoc.pages[0].textFrames[0].fit(FitOptions.FRAME_TO_CONTENT);
			} else if (curDoc.pages.length == 1) {
				var curCont = curDoc.pages[0].textFrames[0].paragraphs;

				for (var i=0;i<curCont.length;i++) {
					if (curCont[i].contents.match("전자파 흡수율")) {
						var count = i;
						break;
					}
				}

				var korTf = curDoc.textFrames.add({
					geometricBounds : [25.0000000018403,15.9999983344793,36,193.999998347582],
				});

				var Conts = curDoc.pages[0].textFrames[1].paragraphs;

				for (var j=count-1; j>=0; j--) {
					Conts[j].select();
					app.cut();
					korTf.insertionPoints[0].select();
					app.paste();
				}

				korTf.textFramePreferences.insetSpacing = [ 4, 4, 4, 4 ];
				korTf.strokeWeight = 0.5;
				korTf.fit(FitOptions.FRAME_TO_CONTENT);
			}
		}
	} catch(e) {
		alert("Error: " + e.line + " : " + e);
		exit();
	}
}