#targetengine "session";

var scriptPath = app.activeScript.path;
// var scriptPath = 'd:/work/220718/_workspace/privateProject/table_utility'
var strokes = app.strokeStyles.everyItem().name;

var w = new Window('palette', 'Table Utility');
w.orientation = 'column';
var all = File(scriptPath + '/all.png');
var lr = File(scriptPath + '/lr.png');
var ud = File(scriptPath + '/ud.png');
var rows = File(scriptPath + '/rows.png');
var columns = File(scriptPath + '/columns.png');
var left = File(scriptPath + '/left.png');
var right = File(scriptPath + '/right.png');

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
var t6 = panelbutton.add('iconbutton', undefined, left, {style: 'toolbutton', toggle:true});
t6.preferredSize = [ 32, 32 ];
var t7 = panelbutton.add('iconbutton', undefined, right, {style: 'toolbutton', toggle:true});
t7.preferredSize = [ 32, 32 ];

var panelopt = w.add ('panel {text: "Settings"}');
panelopt.orientation = 'row';
panelopt.preferredSize = [339, 40];
var opt1 = panelopt.add('statictext', undefined, '선굵기: ');
var strWeight = panelopt.add ('edittext {preferredSize: [30, 20], active: true, text: "0"}');
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
// var panelrun = w.add('panel');
// panelrun.orientation = 'row';
// panelrun.preferredSize = [339, 30]
var btn_apply = w.add('button', [0, 0, 140, 20], '적용하기');
var topt = false

w.show();

t1.onClick = function() {
	t2.value = false;
	t3.value = false;
	t4.value = false;
	t5.value = false;
	t6.value = false;
	t7.value = false;
	topt = true;
}
t2.onClick = function() {
	t1.value = false;
	t3.value = false;
	t4.value = false;
	t5.value = false;
	t6.value = false;
	t7.value = false;
	topt = true;
}
t3.onClick = function() {
	t1.value = false;
	t2.value = false;
	t4.value = false;
	t5.value = false;
	t6.value = false;
	t7.value = false;
	topt = true;
}
t4.onClick = function() {
	t1.value = false;
	t2.value = false;
	t3.value = false;
	t5.value = false;
	t6.value = false;
	t7.value = false;
	topt = true;
}
t5.onClick = function() {
	t1.value = false;
	t2.value = false;
	t3.value = false;
	t4.value = false;
	t6.value = false;
	t7.value = false;
	topt = true;
}
t6.onClick = function() {
	t1.value = false;
	t2.value = false;
	t3.value = false;
	t4.value = false;
	t5.value = false;
	t7.value = false;
	topt = true;
}
t7.onClick = function() {
	t1.value = false;
	t2.value = false;
	t3.value = false;
	t4.value = false;
	t5.value = false;
	t6.value = false;
	topt = true;
}

btn_apply.onClick = function() {
	var docs = app.documents;
	if (docs.length == 0) {
		alert('문서를 열고 실행하세요.');
	} else if (topt == false) {
		alert('Select Border에서 버튼을 클릭하세요.')
	}
	
	if (app.selection[0].constructor.name == 'Table' || app.selection[0].constructor.name == 'Cell') {
		strokeindex = strdrop.selection.index;
		// $.writeln(strWeight.text);
		var tbtype = selectedLineType(panelbutton)
		applytableBorder(tbtype, strWeight.text, strokeindex, myColor.selection.text);
	}
}

function applytableBorder(bordertype, stroke_weight, strokeindex, bordercolor) {
	var doc = app.activeDocument;
	var tableObj = app.selection[0];
	// $.writeln(tableObj.constructor.name);
	// $.writeln(bordertype, strokeindex);
	// $.writeln(bordercolor);
	if (stroke_weight == null || stroke_weight == '0' || stroke_weight == '') {
		strokew = 0;
	} else {
		strokew = stroke_weight;
	}
	var index = Number(strokeindex)
	var bColor = doc.swatches.itemByName(bordercolor);
	// $.writeln(strokew, bordercolor);
	if (bordertype == 'all') {
		if (tableObj.constructor.name == 'Table') {
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
		} else if (tableObj.constructor.name == 'Cell') {
			for (var i=0;i<tableObj.cells.length;i++) {
				var Cells = tableObj.cells[i];
				// stroke weight
				Cells.topEdgeStrokeWeight = strokew;
				Cells.leftEdgeStrokeWeight = strokew;
				Cells.bottomEdgeStrokeWeight = strokew;
				Cells.rightEdgeStrokeWeight = strokew;

				Cells.topEdgeStrokeType = app.strokeStyles[index].name;
				Cells.leftEdgeStrokeType = app.strokeStyles[index].name;
				Cells.bottomEdgeStrokeType = app.strokeStyles[index].name;
				Cells.rightEdgeStrokeType = app.strokeStyles[index].name;
				
				if (bordercolor != null) {
					Cells.topEdgeStrokeColor = bColor;
					Cells.leftEdgeStrokeColor = bColor;
					Cells.bottomEdgeStrokeColor = bColor;
					Cells.rightEdgeStrokeColor = bColor;
				}
			}
		}
	} else if (bordertype == 'lr') {
		if (tableObj.constructor.name == 'Table') {
			tableObj.leftBorderStrokeWeight = strokew;
			tableObj.rightBorderStrokeWeight = strokew;
	
			tableObj.leftBorderStrokeType = app.strokeStyles[index].name;
			tableObj.rightBorderStrokeType = app.strokeStyles[index].name;
			
			if (bordercolor != null) {
				tableObj.leftBorderStrokeColor = bColor;
				tableObj.rightBorderStrokeColor = bColor;
			}
		} else if (tableObj.constructor.name == 'Cell') {
			for (var i=0;i<tableObj.cells.length;i++) {
				var Cells = tableObj.cells[i];
				Cells.leftEdgeStrokeWeight = strokew;
				Cells.rightEdgeStrokeWeight = strokew;
				Cells.leftEdgeStrokeType = app.strokeStyles[index].name;
				Cells.rightEdgeStrokeType = app.strokeStyles[index].name;
				if (bordercolor != null) {
					Cells.leftEdgeStrokeColor = bColor;
					Cells.rightEdgeStrokeColor = bColor;
				}
			}
		}
	} else if (bordertype == 'ud') {
		if (tableObj.constructor.name == 'Table') {
			tableObj.topBorderStrokeWeight = strokew;
			tableObj.bottomBorderStrokeWeight = strokew;
	
			tableObj.topBorderStrokeType = app.strokeStyles[index].name;
			tableObj.bottomBorderStrokeType = app.strokeStyles[index].name;
	
			if (bordercolor != null) {
				tableObj.topBorderStrokeColor = bColor;
				tableObj.bottomBorderStrokeColor = bColor;
			}
		} else if (tableObj.constructor.name == 'Cell') {
			for (var i=0;i<tableObj.cells.length;i++) {
				var Cells = tableObj.cells[i];
				Cells.topEdgeStrokeWeight = strokew;
				Cells.bottomEdgeStrokeWeight = strokew;
				Cells.topEdgeStrokeType = app.strokeStyles[index].name;
				Cells.bottomEdgeStrokeType = app.strokeStyles[index].name;
				if (bordercolor != null) {
					Cells.topEdgeStrokeColor = bColor;
					Cells.bottomEdgeStrokeColor = bColor;
				}
			}
		}
	} else if (bordertype == 'rows') {
		for (var j=0;j<tableObj.cells.length;j++) {
			var Cells = tableObj.cells[j];
			Cells.bottomEdgeStrokeWeight = strokew;
			Cells.bottomEdgeStrokeType = app.strokeStyles[index].name;
			if (bordercolor != null) {
				Cells.bottomEdgeStrokeColor = bColor;
			}
		}
	} else if (bordertype == 'cols') {
		for (var j=0;j<tableObj.cells.length;j++) {
			var Cells = tableObj.cells[j];
			Cells.rightEdgeStrokeWeight = strokew;
			Cells.rightEdgeStrokeType = app.strokeStyles[index].name;
			if (bordercolor != null) {
				Cells.rightEdgeStrokeColor = bColor;
			}
		}
	} else if (bordertype == 'left') {
		for (var j=0;j<tableObj.cells.length;j++) {
			var Cells = tableObj.cells[j];
			Cells.leftEdgeStrokeWeight = strokew;
			Cells.leftEdgeStrokeType = app.strokeStyles[index].name;
			if (bordercolor != null) {
				Cells.leftEdgeStrokeColor = bColor;
			}
		}
	} else if (bordertype == 'right') {
		for (var j=0;j<tableObj.cells.length;j++) {
			var Cells = tableObj.cells[j];
			Cells.rightEdgeStrokeWeight = strokew;
			Cells.rightEdgeStrokeType = app.strokeStyles[index].name;
			if (bordercolor != null) {
				Cells.rightEdgeStrokeColor = bColor;
			}
		}
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
				case 4: return 'cols'
				case 5: return 'left'
				case 6: return 'right'
			}
		}
	}
}
