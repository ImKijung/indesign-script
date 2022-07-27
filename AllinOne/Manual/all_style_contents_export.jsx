main ();  
  
function main() {  
    var curDoc = app.activeDocument;  
    var docPath = curDoc.filePath;   
  
    // create a file  
    var fileToWrite = new File(docPath + "/" + curDoc.name.replace(/\.indd$/,"")  + ".txt");  
    var ok = fileToWrite.open("w");  
        if (!ok) {  
            alert("Error: " + fileToWrite.error);  
            exit();  
    }
    var csv = File (fileToWrite);
    csv.encoding = "UTF-8"; //Encoding for Windows/Excel
   	csv.open ("w");
    // collect the content and write it to the file  
    var allPages= app.activeDocument.pages;
	for (i = 0; i < allPages.length; i++) {
    var allPageTextframes= allPages[i].textFrames;
	for (j = 0; j < allPageTextframes.length; j++)
    {
        var allPageTextframeParagraphs= allPageTextframes[j].paragraphs;
        for (k = 0; k < allPageTextframeParagraphs.length; k++)
        {
            var text = allPageTextframeParagraphs[k].contents;
            var paragraphStyleName = allPageTextframeParagraphs[k].appliedParagraphStyle.name;
           	//$.writeln (paragraphStyleName);
           	//$.writeln (text);
           
				var zeile = paragraphStyleName + "\t" + text;
				csv.writeln (zeile);
         }
     }
 }
  
    fileToWrite.close();  
} // end main  

alert("Finished");  