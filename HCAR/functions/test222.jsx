var doc = app.activeDocument;
var myRef = doc.hyperlinks;
$.writeln(myRef.length);

for (var i=0; i<myRef.length; i++) {
	$.writeln(myRef[i].textSource + ":" + myRef[i].destination.parent.name + "#" + myRef[i].destination.name);
    
}