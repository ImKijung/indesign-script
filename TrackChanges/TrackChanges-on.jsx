var doc = app.activeDocument;
var s = doc.stories;

for (var i=0; i<s.length; i++) {
	s[i].trackChanges = true;
}