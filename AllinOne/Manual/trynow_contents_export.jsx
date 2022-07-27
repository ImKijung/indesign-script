#targetengine "session";

function trynow_contents_export() {
    var w = new Window ("palette","Try Now URL 리스트 만들기");  
    //w.orientation = "hor";
    //var panel_1 = w.add ("panel");
    //	panel_1.add("statictext", undefined, "")

    btn_01 = w.add ("button", [0,0,190,30], "영문: Try Now 리스트");
    btn_02 = w.add ("button", [0,0,190,30], "국문: Try Now 리스트");
    w.show();
        
    btn_01.onClick = function TryNow_ENG () {
        var files;
        var folder = Folder.selectDialog("인디자인이 저장된 폴더를 선택하세요.");
        var trynow_url = new Array();
        var Report = new File(folder +"/" + "indesign_trynow" + ".csv");
        Report.encoding = "ks_c_5601-1987";
        Report.open("w");

        if (folder != null) {
            files = GetFiles(folder);
            if (files.length > 0) {
                // turn off warnings: missing fonts, links, etc.
                app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

            for (z = 0; z < files.length; z++) {
                app.open(files[z]);
                del_icon();
                exportTryNow();
                app.activeDocument.close(SaveOptions.NO);
            } 
            alert("출력완료");
            // turn on warnings
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
            //export_to_Excel()
            Report.close();
            Report.execute();
            }
            else {  
            alert("Found no InDesign documents");
            }
        }

        function exportTryNow() {
            var doc = app.activeDocument;
            app.findGrepPreferences = NothingEnum.nothing; 
            app.changeGrepPreferences = NothingEnum.nothing;
            //app.findGrepPreferences.appliedCharacterStyle = "C_TryNow";

            //Set the find options.
            app.findChangeGrepOptions.includeFootnotes = false;
            app.findChangeGrepOptions.includeHiddenLayers = false;
            app.findChangeGrepOptions.includeLockedLayersForFind = false;
            app.findChangeGrepOptions.includeLockedStoriesForFind = false;
            app.findChangeGrepOptions.includeMasterPages = false;
            
            var greps = [
                {"findWhat":".+Try\\sNow\\r"}
                ]

            for(var i=0; i < greps.length; i++) {
                app.findGrepPreferences.findWhat = greps[i].findWhat;
                var myFound = doc.findGrep();
                for(var k=0; k < myFound.length; k++) {
                    var myContent = myFound[k].contents;
                    var myContent1 = myContent.split('\r');
                    var myContent2 = myContent1.toString().replace(' . ','.');
                    var myContent3 = myContent2.toString().replace('... ','.');
                    var myContent4 = myContent3.toString().replace(' . ','.');
                    var myContent5 = myContent4.toString().replace(' . ','.');
                    var myContent6 = myContent5.toString().replace(' . ','.');
                    var myContent7 = myContent6.toString().replace('..','.');
                    var myContent8 = myContent7.toString().replace(',',' ');
                    var realCont = myContent8.toString().replace(' Try Now','');
                    var myPage = myFound[k].parentTextFrames[0].parentPage.name;
                    var trynow_url = myPage + "," + realCont + "\r";
                    Report.writeln (trynow_url);
                }
            }
        }
    }

    btn_02.onClick = function TryNow_KOR () {
        var files;
        var folder = Folder.selectDialog("인디자인이 저장된 폴더를 선택하세요.");
        var trynow_url = new Array();
        var Report = new File(folder +"/" + "indesign_trynow" + ".csv");
        Report.encoding = "ks_c_5601-1987";
        Report.open("w");

        if (folder != null) {
            files = GetFiles(folder);
            if (files.length > 0) {
                // turn off warnings: missing fonts, links, etc.
                app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

            for (z = 0; z < files.length; z++) {
                app.open(files[z]);
                del_icon();
                exportTryNow();
                app.activeDocument.close(SaveOptions.NO);
            } 
            alert("출력완료");
            // turn on warnings
            app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
            //export_to_Excel()
            Report.close();
            Report.execute();
            }
            else {  
            alert("Found no InDesign documents");
            }
        }

        function exportTryNow() {
            var doc = app.activeDocument;
            app.findGrepPreferences = NothingEnum.nothing; 
            app.changeGrepPreferences = NothingEnum.nothing;
            //app.findGrepPreferences.appliedCharacterStyle = "C_TryNow";

            //Set the find options.
            app.findChangeGrepOptions.includeFootnotes = false;
            app.findChangeGrepOptions.includeHiddenLayers = false;
            app.findChangeGrepOptions.includeLockedLayersForFind = false;
            app.findChangeGrepOptions.includeLockedStoriesForFind = false;
            app.findChangeGrepOptions.includeMasterPages = false;
            
            var greps = [
                {"findWhat":".+실행하기\\r"}
                ]

            for(var i=0; i < greps.length; i++) {
                app.findGrepPreferences.findWhat = greps[i].findWhat;
                var myFound = doc.findGrep();
                for(var k=0; k < myFound.length; k++) {
                    var myContent = myFound[k].contents;
                    var myContent1 = myContent.split('\r');
                    var myContent2 = myContent1.toString().replace(' . ','.');
                    var myContent3 = myContent2.toString().replace('... ','.');
                    var myContent4 = myContent3.toString().replace(' . ','.');
                    var myContent5 = myContent4.toString().replace(' . ','.');
                    var myContent6 = myContent5.toString().replace(' . ','.');
                    var myContent7 = myContent6.toString().replace('..','.');
                    var myContent8 = myContent7.toString().replace(',',' ');
                    var realCont = myContent8.toString().replace(' 실행하기','');
                    var myPage = myFound[k].parentTextFrames[0].parentPage.name;
                    var trynow_url = myPage + "," + realCont + "\r";
                    Report.writeln (trynow_url);
                }
            }
        }
    }

    function del_icon () {
        var doc = app.activeDocument;
        app.findGrepPreferences = NothingEnum.nothing; 
        app.changeGrepPreferences = NothingEnum.nothing;

        app.findChangeGrepOptions.includeFootnotes = false;
        app.findChangeGrepOptions.includeHiddenLayers = false;
        app.findChangeGrepOptions.includeLockedLayersForFind = false;
        app.findChangeGrepOptions.includeLockedStoriesForFind = false;
        app.findChangeGrepOptions.includeMasterPages = false;
        
        app.findGrepPreferences.findWhat = "~a\\s";  
        app.changeGrepPreferences.changeTo = ". ";  
        doc.changeGrep();
    }

    function GetFiles(theFolder) {  
        var files = [],  
        fileList = theFolder.getFiles(),  
        i, file;  
    
        for (m = 0; m < fileList.length; m++) {  
            file = fileList[m];  
            if (file instanceof Folder) {  
                files = files.concat(GetFiles(file));  
            }  
            else if (file instanceof File && file.name.match(/\.indd$/i)) {  
                files.push(file);
            }  
        }  
    
        return files;  
    }
}