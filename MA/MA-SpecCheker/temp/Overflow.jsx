// Overflow - script for finding overflowed text in InDesign
// Pawel Swiecicki 22.08.2006
var of = new Array();
var nodocs = "No documents are opened, so there are no overflows!";
var noframes = "No frames found, so there are no overflows!";
var onpages = "";
var founded = "Overflows found: ";
var oop = "On pages: ";
var notfound = "Uff, no overflows found, Yupi!";

if (app.documents.length < 1){  alert(nodocs);}
else{
  if (app.activeDocument.textFrames.length < 1) {alert(noframes);}
  else
  {
 	 for(var i = 0; i < app.activeDocument.pages.count(); i++)
 	 {
		for(var b = 0; b < app.activeDocument.pages[i].textFrames.count(); b++)
		{ 
		 	if(app.activeDocument.pages[i].textFrames[b].overflows == true)
 		 	{

				of[of.length] = app.activeDocument.pages[i].name;
  		 	}
 		 }
	  }
	  if(of.length > 0)
	  {
	   	for(i = 0; i < of.length; i++)
	  	{
	  		if(i > 0)
	  		{	
	  			onpages += ", ";
	  		}
	  		onpages += of[i]; 
	  	}
 		alert(founded + of.length + "\n\n" + oop + onpages);
	 }
	 else
	 {
	 	alert(notfound);
	 }
 	}
 }