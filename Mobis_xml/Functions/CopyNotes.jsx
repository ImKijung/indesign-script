//반드시 MakeFlatStructure 다음에 호출할 것
//Note(memo)를 xml구조로 복사
function CopyNotesToXML(myDoc) {
    try {
        var stories = myDoc.stories;
        for (var i = 0; i < stories.length; i++) {
            var story = stories[i];
            var notes = story.notes;
            for (var j = 0; j < notes.length; j++) {
                var note = notes[j];
                var contents = '';
                var texts = note.texts;
                for (var k = 0; k < texts.length; k++) {
                    contents += texts[k].contents;
                }
                // var point = note.insertionPoints[0];
                // note.storyOffset.paragraphs[0].characters[-1].select();
                // var mySel = app.selection[0];
                var point = note.storyOffset.paragraphs[0].characters[-1];
                var xml = point.associatedXMLElements[0];
                var attr = xml.xmlAttributes.itemByName('data-note');
                //중복 확인 코드 start
                if (!attr || !attr.isValid)
                    xml.xmlAttributes.add("data-note", contents);
                else
                    attr.value += '\n' + contents;
                //중복 확인 코드 end
            }
        }

        return true;
    } catch (ex) {
        errorMsgs.push('CopyNotesToXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

//Import 한 이후에 호출
//xml 구조에서 Note(memo) 생성
function CopyNotesFromXML(myDoc) {
    RemoveNotes(myDoc);
    try {
        var xmlNotes = myDoc.xmlElements[0].evaluateXPathExpression('//*[@data-note]');
        // $.writeln(xmlNotes.length);
        for (var i = 0; i < xmlNotes.length; i++) {
            var xmlNote = xmlNotes[i];
            // app.select(xmlNote);
            var hrefstring = xmlNote.xmlAttributes.itemByName('data-note').value;
            var hrefs = hrefstring.split('\n');
            for (var j = 0; j < hrefs.length ; j++) {
                var myNote = xmlNote.texts[-1].insertionPoints[-2].notes.add();
                myNote.texts[0].contents = hrefs[j];
                // xmlNote.texts[-1].insertionPoints[-2].contents = " ";
            }
        }
        return true;
    } catch (ex) {
        errorMsgs.push('CopyNotesFromXML 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}

//기존 Note(memo) 제거
function RemoveNotes(myDoc) {
    try {
        var stories = myDoc.stories;
        for (var i = 0; i < stories.length; i++) {
            var story = stories[i];
            var notes = story.notes;
            while (notes.length > 0) {
                notes[0].remove();
                notes = story.notes;
            }
        }
        return true;
    } catch (ex) {
        errorMsgs.push('RemoveNotes 함수 오류(' + myDoc.name + '): Line:' + ex.line + ':: ' + ex);
        hasError = true;
        throw ex;
    }
}