#targetengine "session";

var scriptPath = 'C:/Users/adfasdfasdf/AppData/Roaming/Adobe/InDesign/Version 16.0-J/ko_KR/Scripts/Scripts Panel/TEST/kkk/'
var strokes = app.strokeStyles.everyItem().name;

var w = new Window('palette', 'Table Utility');
w.orientation = 'column';
var all = File(scriptPath + 'all.png');
var lr = File(scriptPath + 'lr.png');
var ud = File(scriptPath + 'ud.png');
var rows = File(scriptPath + 'rows.png');
var columns = File(scriptPath + 'columns.png');

var panelbutton = w.add ('panel {text: "Selet Border"}');
panelbutton.preferredSize = [339, 40];
panelbutton.orientation = 'row';
var t1 = panelbutton.add('iconbutton', undefined, all, {style: 'toolbutton', toggle:true});
t1.preferredSize = [ 32, 32 ];
var t2 = panelbutton.add('iconbutton', undefined, lr, {style: 'toolbutton', toggle:true});
t2.preferredSize = [ 32, 32 ];
var t3 = panelbutton.add('iconbutton', undefined, ud, {style: 'toolbutton', toggle:true});
t3.preferredSize = [ 32, 32 ];
var t4 = panelbutton.add('iconbutton', undefined, rows, {style: 'toolbutton', toggle:true});
t4.preferredSize = [ 32, 32 ];
var t5 = panelbutton.add('iconbutton', undefined, columns, {style: 'toolbutton', toggle:true});
t5.preferredSize = [ 32, 32 ];

var panelopt = w.add ('panel {text: "Settings"}');
panelopt.orientation = 'row';
var opt1 = panelopt.add('statictext', undefined, '선굵기: ');
var strWeight = panelopt.add ('edittext {preferredSize: [30, 20], active: true}');
var opt2 = panelopt.add('statictext', undefined, '선유형: ');
var strdrop = panelopt.add('dropdownlist', [0, 0, 50, 20]);
// strdrop.add('item', strokes);
for (var i=0;i<strokes.length;i++) {
	strdrop.add('item', strokes[i]);
}
strdrop.selection = 17;
var opt3 = panelopt.add('statictext', undefined, '선색상: ');
var myColor = panelopt.add('dropdownlist', [0, 0, 50, 20]);
var swats = app.swatches
for (var i=0;i<swats.length;i++) {
	myColor.add('item', swats[i].name);
}
var panelrun = w.add('panel');
panelrun.orientation = 'row';
panelrun.preferredSize = [339, 30]
var btn_apply = panelrun.add('button', [0, 0, 140, 20], '적용하기');
var topt = false

w.show();

t1.onClick = function() {
	t2.value = false;
	t3.value = false;
	t4.value = false;
	t5.value = false;
	topt = true;
}

t2.onClick = function() {
	t1.value = false;
	t3.value = false;
	t4.value = false;
	t5.value = false;
	topt = true;
}

t3.onClick = function() {
	t1.value = false;
	t2.value = false;
	t4.value = false;
	t5.value = false;
	topt = true;
}

t4.onClick = function() {
	t1.value = false;
	t2.value = false;
	t3.value = false;
	t5.value = false;
	topt = true;
}

t5.onClick = function() {
	t1.value = false;
	t2.value = false;
	t3.value = false;
	t4.value = false;
	topt = true;
}

btn_apply.onClick = function() {
	var docs = app.documents;
	if (docs.length == 0) {
		alert('문서를 열고 실행하세요.');
	}
	if (app.selection[0] == null) {
		alert('표를 선택한 다음 실행하세요.')
	}
	if (topt == false) {
		alert('Select Border에서 버튼을 클릭하세요.')
	}
	
	if (app.selection[0].constructor.name == 'Table') {
		strokeindex = strdrop.selection.index;
		// $.writeln(strokeindex);
		var tbtype = selectedLineType(panelbutton)
		applytableBorder(tbtype, strWeight, strokeindex, myColor.selection.text);
	} else if (app.selection[0].constructor.name == 'Cell') {
	
	}
}

function applytableBorder(bordertype, strWeight, strokeindex, bordercolor) {
	var doc = app.activeDocument;
	var tableObj = app.selection[0];
	// $.writeln(bordertype, strokeindex);
	
	if (strWeight == null && strWeight == '0') {
		strokew = 0;
	} else {
		strokew = Number(strWeight.text);
	}
	var index = Number(strokeindex)
	var bColor = doc.swatches.itemByName(bordercolor);
	$.writeln(bordercolor);
	if (bordertype == 'all') {
		// stroke weight
		tableObj.topBorderStrokeWeight = strokew;
		tableObj.leftBorderStrokeWeight = strokew;
		tableObj.bottomBorderStrokeWeight = strokew;
		tableObj.rightBorderStrokeWeight = strokew;
		
		// stroke type
		tableObj.topBorderStrokeType = app.strokeStyles[index].name;
		tableObj.leftBorderStrokeType = app.strokeStyles[index].name;
		tableObj.bottomBorderStrokeType = app.strokeStyles[index].name;
		tableObj.rightBorderStrokeType = app.strokeStyles[index].name;

		if (bordercolor != null) {
			tableObj.topBorderStrokeColor = bColor;
			tableObj.leftBorderStrokeColor = bColor;
			tableObj.bottomBorderStrokeColor = bColor;
			tableObj.rightBorderStrokeColor = bColor;
		}
	} else if (bordertype == 'lr') {
		tableObj.leftBorderStrokeWeight = strokew;
		tableObj.rightBorderStrokeWeight = strokew;

		tableObj.leftBorderStrokeType = app.strokeStyles[index].name;
		tableObj.rightBorderStrokeType = app.strokeStyles[index].name;
		
		if (bordercolor != null) {
			tableObj.leftBorderStrokeColor = bColor;
			tableObj.rightBorderStrokeColor = bColor;
		}
	} else if (bordertype == 'ud') {
		tableObj.topBorderStrokeWeight = strokew;
		tableObj.bottomBorderStrokeWeight = strokew;

		tableObj.topBorderStrokeType = app.strokeStyles[index].name;
		tableObj.bottomBorderStrokeType = app.strokeStyles[index].name;

		if (bordercolor != null) {
			tableObj.topBorderStrokeColor = bColor;
			tableObj.bottomBorderStrokeColor = bColor;
		}
	} else if (bordertype == 'rows') {
		
	}
	
	
}

function selectedLineType(tbuttons) {
	for (var x=0; x<tbuttons.children.length; x++) {
		if (tbuttons.children[x].value == true) {
			switch(x) {
				case 0: return 'all'
				case 1: return 'lr'
				case 2: return 'ud'
				case 3: return 'rows'
				case 4: return 'colse'
			}
		}
	}
}