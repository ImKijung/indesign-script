var doc = app.activeDocument;

var myObj = app.selection[0];

app.copy();

var bt = new BridgeTalk();
bt.target = 'illustrator';
bt.body = 'app.paste();';
bt.body += 'app.doScript("expandAll", "expand", true);';
bt.body += 'app.unloadAction(expand, "");';
bt.body += '$.sleep(1000);';
// bt.body += 'app.cut()';
bt.onResult = function(resObj) {
	var myResult = resObj.body;
}
bt.onError = function( inBT ) { alert(inBT.body); };
bt.send(8);

// app.paste();