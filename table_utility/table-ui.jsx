#targetengine "session";

var scriptPath = app.activeScript.path;
// var scriptPath = 'd:/work/220718/_workspace/privateProject/table_utility'

if (app.documents.length == 0) {
	alert('문서를 열어 놓은 후 실행하세요.');
	exit();
}
var doc = app.activeDocument;
var table_color01 = {
	name: 'table-color1',
	colorValue : [20, 0, 0, 60],
	space: ColorSpace.CMYK
}
var table_color02 = {
	name: 'table-color2',
	colorValue : [0, 20, 0, 60],
	space: ColorSpace.CMYK
}
var table_color03 = {
	name: 'table-color3',
	colorValue : [0, 0, 20, 60],
	space: ColorSpace.CMYK
}
var myColors = [table_color01, table_color02, table_color03];
for (var x=0;x<myColors.length;x++) {
	var tableColor = doc.colors.itemByName(myColors[x].name);
	if (!tableColor.isValid) {
		doc.colors.add(myColors[x]);
	}
}

var w = new Window('palette', 'Table Utility');
w.orientation = 'row';
var all = File(scriptPath + '/all.png');
var lr = File(scriptPath + '/lr.png');
var ud = File(scriptPath + '/ud.png');
var every = File(scriptPath + '/every.png');
// var columns = File(scriptPath + '/columns.png');
// var left = File(scriptPath + '/left.png');
// var right = File(scriptPath + '/right.png');

var panelbutton = w.add ('panel {text: "Selet Border"}');
panelbutton.preferredSize = [100, 190];
panelbutton.orientation = 'column';
var t1 = panelbutton.add('iconbutton', undefined, all, {style: 'toolbutton', toggle:true});
t1.preferredSize = [ 32, 32 ];
var t2 = panelbutton.add('iconbutton', undefined, lr, {style: 'toolbutton', toggle:true});
t2.preferredSize = [ 32, 32 ];
var t3 = panelbutton.add('iconbutton', undefined, ud, {style: 'toolbutton', toggle:true});
t3.preferredSize = [ 32, 32 ];
var t4 = panelbutton.add('iconbutton', undefined, every, {style: 'toolbutton', toggle:true});
t4.preferredSize = [ 32, 32 ];
// var t5 = panelbutton.add('iconbutton', undefined, columns, {style: 'toolbutton', toggle:true});
// t5.preferredSize = [ 32, 32 ];
// var t6 = panelbutton.add('iconbutton', undefined, left, {style: 'toolbutton', toggle:true});
// t6.preferredSize = [ 32, 32 ];
// var t7 = panelbutton.add('iconbutton', undefined, right, {style: 'toolbutton', toggle:true});
// t7.preferredSize = [ 32, 32 ];

var panelopt = w.add ('panel {text: "Settings"}');
panelopt.orientation = 'column';
panelopt.preferredSize = [175, 190];
var g1 = panelopt.add('group {orientation: "row", alignChildren: "left"}')
var opt1 = g1.add('statictext', undefined, '선굵기: ');
var strWeight = g1.add ('edittext {preferredSize: [90, 20], active: true, text: "0"}');

var g2 = panelopt.add('group {orientation: "row", alignChildren: "left"}')
var opt2 = g2.add('statictext', undefined, '선유형: ');
var strdrop = g2.add('dropdownlist', [0, 0, 90, 20], ['실선', '일본식 점선', '파선 3:2']);
strdrop.selection = 0;

var g3 = panelopt.add('group {orientation: "row", alignChildren: "left"}')
var opt3 = g3.add('statictext', undefined, '선색상: ');
var sColor = g3.add('dropdownlist', [0, 0, 90, 20], ['table-color1', 'table-color2', 'table-color3']);
sColor.selection = 0;
var chkbox = panelopt.add('checkbox', undefined, '문서 내 모든 테이블 설정');
var btn_apply = panelopt.add('button', undefined, '적용하기');
var topt = false

w.show();

t1.onClick = function() {
	t2.value = false;
	t3.value = false;
	t4.value = false;
	topt = true;
}
t2.onClick = function() {
	t1.value = false;
	t3.value = false;
	t4.value = false;
	topt = true;
}
t3.onClick = function() {
	t1.value = false;
	t2.value = false;
	t4.value = false;
	topt = true;
}
t4.onClick = function() {
	t1.value = false;
	t2.value = false;
	t3.value = false;
	topt = true;
}

btn_apply.onClick = function() {
	var docs = app.documents;
	if (docs.length == 0) {
		alert('문서를 열고 실행하세요.');
	} else if (topt == false) {
		alert('Select Border에서 버튼을 클릭하세요.')
	}
	
	if ((app.selection[0].constructor.name == 'Table' || app.selection[0].constructor.name == 'Cell') && (chkbox.value == false)) {
		strokeName = strdrop.selection.text;
		var tbtype = selectedLineType(panelbutton)
		applytableBorder(tbtype, strWeight.text, strokeName, sColor.selection.text);
	} else {
		if (chkbox.value == true) {
			strokeName = strdrop.selection.text;
			var tbtype = selectedLineType(panelbutton)
			applyalltables(tbtype, strWeight.text, strokeName, sColor.selection.text);
		}
	}
}

function applyalltables(bordertype, stroke_weight, strokeName, colorName) {
	var doc = app.activeDocument;
	if (stroke_weight == null || stroke_weight == '0' || stroke_weight == '') {
		strokew = 0;
	} else {
		strokew = stroke_weight;
	}
	var bColor = doc.swatches.itemByName(colorName);
	var myStoryes = doc.stories;
	try {
		for (var j=0; j<myStoryes.length; j++) {
			var myTables = myStoryes[j].tables;
			for (var i=0;i<myTables.length;i++) {
				var tableObj = myTables[i];
				if (bordertype == 'all' || bordertype == 'every') {
					tableObj.topBorderStrokeWeight = strokew;
					tableObj.leftBorderStrokeWeight = strokew;
					tableObj.bottomBorderStrokeWeight = strokew;
					tableObj.rightBorderStrokeWeight = strokew;
					tableObj.topBorderStrokeType = strokeName;
					tableObj.leftBorderStrokeType = strokeName;
					tableObj.bottomBorderStrokeType = strokeName;
					tableObj.rightBorderStrokeType = strokeName;
					tableObj.topBorderStrokeColor = bColor;
					tableObj.leftBorderStrokeColor = bColor;
					tableObj.bottomBorderStrokeColor = bColor;
					tableObj.rightBorderStrokeColor = bColor;
					if (bordertype == 'every') {
						for (var n=0; n<tableObj.cells.length; n++) {
							var nCells = tableObj.cells[n];
							nCells.topEdgeStrokeWeight = strokew;
							nCells.leftEdgeStrokeWeight = strokew;
							nCells.bottomEdgeStrokeWeight = strokew;
							nCells.rightEdgeStrokeWeight = strokew;
							nCells.topEdgeStrokeType = strokeName;
							nCells.leftEdgeStrokeType = strokeName;
							nCells.bottomEdgeStrokeType = strokeName;
							nCells.rightEdgeStrokeType = strokeName;
							nCells.topEdgeStrokeColor = bColor;
							nCells.leftEdgeStrokeColor = bColor;
							nCells.bottomEdgeStrokeColor = bColor;
							nCells.rightEdgeStrokeColor = bColor;
						}
					}
				} else if (bordertype == 'lr') {
					tableObj.leftBorderStrokeWeight = strokew;
					tableObj.rightBorderStrokeWeight = strokew;
					tableObj.leftBorderStrokeType = strokeName;
					tableObj.rightBorderStrokeType = strokeName;
					tableObj.leftBorderStrokeColor = bColor;
					tableObj.rightBorderStrokeColor = bColor;
				} else if (bordertype == 'ud') {
					tableObj.topBorderStrokeWeight = strokew;
					tableObj.bottomBorderStrokeWeight = strokew;
					tableObj.topBorderStrokeType = strokeName;
					tableObj.bottomBorderStrokeType = strokeName;
					tableObj.topBorderStrokeColor = bColor;
					tableObj.bottomBorderStrokeColor = bColor;
				}
			}
		}
	} catch(e) {
		alert(e.line + ":" + e);
	}
}

function applytableBorder(bordertype, stroke_weight, strokeName, colorName) {
	var doc = app.activeDocument;
	var tableObj = app.selection[0];
	// $.writeln(tableObj.constructor.name);
	// $.writeln(bordertype, strokeindex);
	// $.writeln(strokeName);
	if (stroke_weight == null || stroke_weight == '0' || stroke_weight == '') {
		strokew = 0;
	} else {
		strokew = stroke_weight;
	}
	var bColor = doc.swatches.itemByName(colorName);
	try {
		if (bordertype == 'all' || bordertype == 'every') {
			if (tableObj.constructor.name == 'Table') {
				// stroke weight
				tableObj.topBorderStrokeWeight = strokew;
				tableObj.leftBorderStrokeWeight = strokew;
				tableObj.bottomBorderStrokeWeight = strokew;
				tableObj.rightBorderStrokeWeight = strokew;
				// stroke type
				tableObj.topBorderStrokeType = strokeName;
				tableObj.leftBorderStrokeType = strokeName;
				tableObj.bottomBorderStrokeType = strokeName;
				tableObj.rightBorderStrokeType = strokeName;
				tableObj.topBorderStrokeColor = bColor;
				tableObj.leftBorderStrokeColor = bColor;
				tableObj.bottomBorderStrokeColor = bColor;
				tableObj.rightBorderStrokeColor = bColor;
				if (bordertype == 'every') {
					for (var n=0; n<tableObj.cells.length; n++) {
						var nCells = tableObj.cells[n];
						nCells.topEdgeStrokeWeight = strokew;
						nCells.leftEdgeStrokeWeight = strokew;
						nCells.bottomEdgeStrokeWeight = strokew;
						nCells.rightEdgeStrokeWeight = strokew;
						nCells.topEdgeStrokeType = strokeName;
						nCells.leftEdgeStrokeType = strokeName;
						nCells.bottomEdgeStrokeType = strokeName;
						nCells.rightEdgeStrokeType = strokeName;
						nCells.topEdgeStrokeColor = bColor;
						nCells.leftEdgeStrokeColor = bColor;
						nCells.bottomEdgeStrokeColor = bColor;
						nCells.rightEdgeStrokeColor = bColor;
					}
				}
			} else if (tableObj.constructor.name == 'Cell') {
				tableObj.parent.topBorderStrokeWeight = strokew;
				tableObj.parent.leftBorderStrokeWeight = strokew;
				tableObj.parent.bottomBorderStrokeWeight = strokew;
				tableObj.parent.rightBorderStrokeWeight = strokew;
				tableObj.parent.topBorderStrokeType = strokeName;
				tableObj.parent.leftBorderStrokeType = strokeName;
				tableObj.parent.bottomBorderStrokeType = strokeName;
				tableObj.parent.rightBorderStrokeType = strokeName;
				tableObj.parent.topBorderStrokeColor = bColor;
				tableObj.parent.leftBorderStrokeColor = bColor;
				tableObj.parent.bottomBorderStrokeColor = bColor;
				tableObj.parent.rightBorderStrokeColor = bColor;
			}
		} else if (bordertype == 'lr') {
			if (tableObj.constructor.name == 'Table') {
				tableObj.leftBorderStrokeWeight = strokew;
				tableObj.rightBorderStrokeWeight = strokew;
				tableObj.leftBorderStrokeType = strokeName;
				tableObj.rightBorderStrokeType = strokeName;
				tableObj.leftBorderStrokeColor = bColor;
				tableObj.rightBorderStrokeColor = bColor;
			} else if (tableObj.constructor.name == 'Cell') {
				tableObj.parent.leftBorderStrokeWeight = strokew;
				tableObj.parent.rightBorderStrokeWeight = strokew;
				tableObj.parent.leftBorderStrokeType = strokeName;
				tableObj.parent.rightBorderStrokeType = strokeName;
				tableObj.parent.leftBorderStrokeColor = bColor;
				tableObj.parent.bottomBorderStrokeColor = bColor;
			}
		} else if (bordertype == 'ud') {
			if (tableObj.constructor.name == 'Table') {
				tableObj.topBorderStrokeWeight = strokew;
				tableObj.bottomBorderStrokeWeight = strokew;
				tableObj.topBorderStrokeType = strokeName;
				tableObj.bottomBorderStrokeType = strokeName;
				tableObj.topBorderStrokeColor = bColor;
				tableObj.bottomBorderStrokeColor = bColor;
			} else if (tableObj.constructor.name == 'Cell') {
				tableObj.parent.topBorderStrokeWeight = strokew;
				tableObj.parent.bottomBorderStrokeWeight = strokew;
				tableObj.parent.topBorderStrokeType = strokeName;
				tableObj.parent.bottomBorderStrokeType = strokeName;
				tableObj.parent.topBorderStrokeColor = bColor;
				tableObj.parent.bottomBorderStrokeColor = bColor;
			}
		}
	} catch(e) {
		alert(e.line + ":" + e);
	}
}

function selectedLineType(tbuttons) {
	for (var x=0; x<tbuttons.children.length; x++) {
		if (tbuttons.children[x].value == true) {
			switch(x) {
				case 0: return 'all'
				case 1: return 'lr'
				case 2: return 'ud'
				case 3: return 'every'
				// case 4: return 'cols'
				// case 5: return 'left'
				// case 6: return 'right'
			}
		}
	}
}
