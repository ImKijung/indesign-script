try {
var myMenuAction = app.menuActions.itemByID (1512546).checked
}
catch (e) {alert("PV 마킹이 꺼져있어")}

if (myMenuAction == true) {
	try {
	app.menuActions.itemByID (1512544).invoke();
	}
	catch (e) {}
}
if (myMenuAction == false) {
	try {
	app.menuActions.itemByID (1512546).invoke();
	}
	catch (e) {}
}