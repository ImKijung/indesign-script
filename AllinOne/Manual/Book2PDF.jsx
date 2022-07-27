var myFolder = Folder.selectDialog ("북파일이 있는 폴더를 선택하세요.");
var myFileIn = File(myFolder).getFiles("*.indb");
var myFileOut = Folder.selectDialog("PDF가 저장될 폴더를 선택하세요.");
var params = new Array();
params = getData( myFolder, myFileIn );
app.dialogs.everyItem().destroy();
Opens();

 

function getData(PDF)
{
var pdfPresets = app.pdfExportPresets.everyItem().name;
var dlg = app.dialogs.add( { name : 'Choose Style' } );
with(dlg)
{
with( dialogColumns.add() )
{
with( borderPanels.add( ) )
{
staticTexts.add( { staticLabel : 'PDF Presets:', minWidth : 93 } );
dropDown = dropdowns.add( { stringList : pdfPresets, selectedIndex : 0 });
}
}
           }
if( dlg.show() == false ){
dlg.destroy();
exit(0);
}
return dropDown.selectedIndex;
}

 

function Opens()
{
for(var i=0; i<myFileIn.length; i++) {
    var myDoc =app.open(myFileIn[i], true);
    var Final = new Array();
    var myFileInPath = myDoc.filePath;
    var PDFs = new Folder(myFileOut);
    pdfUsing (params, PDFs, '.pdf');
    myDoc.close(SaveOptions.NO);
    }
}

 

function pdfUsing(presetName, folderPath, fileExt)
{
   var targFolder = new Folder(folderPath);
    try{
		myFile = new File(targFolder.fsName + "/" + app.activeBook.name.split(".indb")[0] + fileExt);
		app.activeBook.exportFile(ExportFormat.pdfType, myFile, false, app.pdfExportPresets.item(presetName));
}catch(e){
alert('Error in pdfUsing: ' + e)
  }
}

alert ("Finished");

//--------------------------------