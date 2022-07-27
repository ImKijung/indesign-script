var doc = app.activeDocument; 
var unassign = doc.conditions.item("mmi-unassign");
app.findTextPreferences = app.changeTextPreferences = null;
app.findTextPreferences.appliedConditions = [ unassign ];
var myFound = doc.findText();
if (myFound.length == 0) unassign.remove();
app.findTextPreferences = app.changeTextPreferences = null;