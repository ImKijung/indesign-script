#targetengine "session";

function index_contents_book_export() {
    if (app.books.length == 0) {
        alert("인디자인 북 파일을 열어주세요.");
    } else if (app.books.length > 0) {
        book_export();
    }
}

function book_export() {
    var myBook = app.activeBook;
    var myBookContents = myBook.bookContents.everyItem().getElements();
    var folder = myBook.filePath;
    var myPrepend = prompt("문자 스타일명을 입력하세요.", "", "문자 스타일 목록 추출하기");  
    if (!myPrepend) exit(); 
    if (myPrepend) {
        // var index_content = new Array();
        var Report = new File(folder +"/" + myPrepend + ".txt");
        Report.open("w");
        Report.encoding = "utf-8";
        // turn off warnings: missing fonts, links, etc.
        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
        for(var n=0; n<myBookContents.length; n++) {
            var myPath = myBookContents[n].fullName;
            var myFile = File(myPath);
            var doc = app.open(myFile);
            main (doc);
            doc.close(SaveOptions.NO);
        }
    function main() {
        var myErr = mySucc = 0;
        var myList = "";
        //var myCharacterStyle = "Iword";

        var doc = app.activeDocument;
        if (doc.characterStyles.item(myPrepend) == null) {
            
        } else {
            app.changeTextPreferences = NothingEnum.nothing;    
            app.findTextPreferences = NothingEnum.nothing;    
            app.findTextPreferences.appliedCharacterStyle = myPrepend;
            var myFound = app.activeDocument.findText();

            for(var k=0; k < myFound.length; k++) {
                var myContent = myFound[k].contents;
                myContent = myContent.replace(/\ufeff/g, "");
                var realCont = myContent.toString().split("\r");
                //var realCont = myContent1.replace('~a', 'icon')
                var myPage = myFound[k].parentTextFrames[0].parentPage.name;
                var index_content = myPage + "\t" + realCont + "\r";
                Report.writeln (index_content);
                }
            }
        }
        app.changeTextPreferences = NothingEnum.nothing;
        app.findTextPreferences = NothingEnum.nothing;
        //export_to_Excel()
    }
    // turn on warnings
    app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
    alert("출력완료");
    Report.close();
    Report.execute();
}