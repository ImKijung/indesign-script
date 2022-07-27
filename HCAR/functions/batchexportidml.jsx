function batchExportIDML() {
     //app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
     var myFolder = Folder.selectDialog( "인디자인 문서가 있는 상위 폴더를 선택하세요." );  
     if ( myFolder != null ) {
          var myFiles = [];
          GetSubFolders(myFolder);
          if ( myFiles.length > 0 ) {
          // turn off warnings: missing fonts, links, etc.
          app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
          var myBookFileName = myFolder + "/"+ myFolder.name;
          myBookFile = new File( myBookFileName );
               if ( myBookFile.exists ) {  
                    if ( app.books.item(myFolder.displayName) == null ) {
                         myBook = app.open( myBookFile );
                    }
               }
               var myWindow = new Window ('palette', "작업 진행 중 ...");
			     myWindow.pbar = myWindow.add ('progressbar', undefined, 0, myFiles.length);
			     myWindow.pbar.preferredSize.width = 300;
		     myWindow.show();
               for ( i=0; i < myFiles.length; i++ ) {
                    myWindow.pbar.value = i+1;
                    var myBookFile = app.open(File(myFiles[i]), true);
                    with ( myBookFile ) {
                         exportIDML(myBookFile)
                    }
               }
               myWindow.close();
          // turn on warnings
          app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;  
          alert ("모든 문서에 IDML 파일을 추출했습니다.");
          }
     }

function exportIDML(myBookFile) {
     myFilePath = myBookFile.filePath.absoluteURI + "/" + myBookFile.name.split(".indd")[0] + ".idml"; //파일명 설정
     myFile = new File(myFilePath);
     myBookFile.exportFile(ExportFormat.INDESIGN_MARKUP, myFile); 
     myBookFile.close(SaveOptions.no)
}

function UpdateAllOutdatedLinks(myBookFile) {
     for (var d = myBookFile.links.length-1; d >= 0; d--) {
          var link = myBookFile.links[d];
          if (link.status == LinkStatus.linkOutOfDate) {
               link.update();
          }
     }
}

//=================================== FUNCTIONS =========================================  
function GetSubFolders(theFolder) {
     var myFileList = theFolder.getFiles();
     for (var i = 0; i < myFileList.length; i++) {
          var myFile = myFileList[i];
          if (myFile instanceof Folder) {
               GetSubFolders(myFile);
          }
          else if (myFile instanceof File && myFile.name.match(/\.indd$/i)) {
               myFiles.push(myFile);
          }
     }
}
//-------------------------------------------------------------------------------------------------------------- 
}