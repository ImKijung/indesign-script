function index_contents_docu_export() {
    if (app.documents.length == 0) {  
        alert("인디자인 문서를 열고 실행해주세요.");  
        exit();
    } else docu_export();
}

function docu_export() {
    var myErr = mySucc = 0;    
    var myList = "";  
    var doc = app.activeDocument      
    var myCharacterStyle = myDisplayDialog(doc);    
    // var Index_txt = new Array()
    var docPath = doc.filePath;  

    var Report = new File(docPath +"/" + doc.name.replace(".indd","") + "_" + myCharacterStyle + ".txt");
    Report.open("w");
    Report.encoding = "utf-8";

    app.changeTextPreferences = NothingEnum.nothing;    
    app.findTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.appliedCharacterStyle = myCharacterStyle;       

    var myFound = doc.findText();
        for(i=0; i<myFound.length; i++){
        //app.select(myFound[i])
        //alert(myFound[i].contents)
            var myContent = myFound[i].contents;
            myContent = myContent.replace(/\ufeff/g, "");
            //var myCont = myContent.split("\x{FEFF}");
            var myPage = myFound[i].parentTextFrames[0].parentPage.name;
            Report.writeln(myPage + "\t" + myContent + "\r")
        }
    //alert(Index_txt)

    // Report.write(Index_txt);
    Report.close();
    Report.execute();
    
    
    function myDisplayDialog(doc) {     
        var myFieldWidth = 120;     
        
        var myCharStyles = doc.characterStyles.everyItem().name;     
        
        var myDialog = app.dialogs.add({name:"문자스타일 콘텐츠 추출"});     
        with(myDialog.dialogColumns.add()){     
            with(dialogRows.add()){     
                with(dialogColumns.add()){     
                    staticTexts.add({staticLabel:"Character style:", minWidth:myFieldWidth});     
                }     
                with(dialogColumns.add()){     
                    var mySourceDropdown = dropdowns.add({stringList:myCharStyles, selectedIndex:myCharStyles.length-1});     
                }     
            }     
        }     
        var myResult = myDialog.show();     
        if(myResult == true){     
            var theCharStyle =myCharStyles[mySourceDropdown.selectedIndex];     
            myDialog.destroy();     
        }     
        else{     
            myDialog.destroy()     
            exit();     
        }     
        return theCharStyle;     
    }    
   
    function selectIt( theObj ) {       
        var myZoom = 400;  
        app.select(theObj,SelectionOptions.replaceWith);    
        app.activeWindow.zoomPercentage = myZoom;    
        
        // Option to terminate if called within a loop    
        var myChoice = confirm ("The reference is buggy!\rMore?");     
        if (myChoice == false)     
            exit();     
        return app.selection[0];   
    } 
}