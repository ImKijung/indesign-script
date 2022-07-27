var doc = app.activeDocument;

var tables = doc.xmlElements[0].evaluateXPathExpression("descendant::Table[not(@HeaderRowCount)]");
var tables_count = tables.length;
if (tables_count > 0) {
	for (var i=0 ; i<tables_count ; i++) {
		var table = tables[i].tables[0];
		tables[i].xmlAttributes.add('HeaderRowCount', table.headerRowCount + '');
		var rows = table.rows;
		var rows_count = rows.length;
		for (var j=0 ; j<rows_count ; j++)
			rows[j].cells[0].associatedXMLElement.xmlAttributes.add('newrow', 'newrow');
		var columns = table.columns;
		var columns_count = columns.length;
		var widths = [];
		for (var j=0 ; j<columns_count ; j++)
			widths = widths.concat((columns[j].width/table.width).toFixed(2) + '*');
			tables[i].xmlAttributes.add('colspecs', widths.join(':'));
	}
	var cells = doc.xmlElements[0].evaluateXPathExpression("descendant::Cell[not(@namest)]");
	var cells_count = cells.length;
	for (var i=0 ; i<cells_count ; i++) {
		var cell = cells[i].cells[0];
		if (cell.columnSpan > 1) {
			var namest = cell.parentColumn.index + 1;
			var nameend = namest + cell.columnSpan - 1;
			cells[i].xmlAttributes.add('namest', 'col' + namest);
			cells[i].xmlAttributes.add('nameend', 'col' + nameend);
		}
	}
}