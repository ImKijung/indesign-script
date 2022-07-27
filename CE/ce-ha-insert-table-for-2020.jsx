// var Excelfile = "D:\\TCS\\_develop\\067_CE-HA_Excel-import-Script\\original-merge-cell.xlsm_2020-07-29.xlsx";
// D:\\TCS\\_develop\\067_CE-HA_Excel-import-Script\\original-merge-cell.xlsm_2020-07-29.xlsx
// d:\\TCS\\_develop\\067_CE-HA_Excel-import-Script\\original-merge-cell.xlsm_2020-07-29.xlsx

var Excelfile = File.openDialog("엑셀 파일을 선택하세요");
// var myPath = Excelfile.toString();
// var xvbsPath = myPath.replace(/\//g, "\\\\");
// var vbsPath = xvbsPath.replace(/^\\\\(\w)/g, "$1:");
// $.writeln(vbsPath);
// var shtNumber = GetDataFromExcelPC(vbsPath)
// var shtcount = shtNumber.toString().replace(';','');
// shtcount = shtcount -1;
// var shtcount = 3;
// $.writeln(shtcount);
var doc = app.activeDocument;

var myProperties = app.excelImportPreferences.properties = {
	tableFormatting : TableFormattingOptions.EXCEL_UNFORMATTED_TABLE,
	useTypographersQuotes : true,
	preserveGraphics : true,
	decimalPlaces : 2,
	showHiddenCells : false,
	errorCode : 0
}
doc.pages.item(-1).textFrames[0].place(Excelfile, true, myProperties);
doc.pages.add();
// var textFrame = doc.pages[-1].textFrames.add();
// 	textFrame.properties = {
// 		geometricBounds: [14,15.000,275.999,191.999]
// 	}
alert("Complete");

// for (var i=0; i<shtcount; i++) {
// 	var myProperties = app.excelImportPreferences.properties = {
// 		tableFormatting : TableFormattingOptions.EXCEL_UNFORMATTED_TABLE,
// 		useTypographersQuotes : true,
// 		preserveGraphics : true,
//		sheetindex : 0,
// 		decimalPlaces : 2,
// 		showHiddenCells : false,
// 		errorCode : 0
// 	}
// 	doc.pages.item(-1).textFrames[0].place(Excelfile, false, myProperties);
// 	doc.pages.add();
// 	var textFrame = doc.pages[-1].textFrames.add();
// 		textFrame.properties = {
// 			geometricBounds: [14,15.000,275.999,191.999]
// 		}
// }

// var allTables = doc.stories.everyItem().tables.everyItem().getElements();
// var tbStyle = doc.tableStyles.itemByName("Common-Bordered");
// for (var i=0;i<allTables.length;i++) {
// 	var myTable = allTables[i];
// 	myTable.appliedTableStyle = tbStyle;
// 	myTable.width = "177mm";
// }


// function GetDataFromExcelPC(excelFilePath) {
// 	var vbs = 'Public excelFilePath\r';
// 	vbs += 'Function ReadFromExcel()\r';
// 	vbs += 'Dim worksheetcount\r';
// 	vbs += 'Set objExcel = CreateObject("Excel.Application")\r';
// 	vbs += 'Set objBook = objExcel.Workbooks.Open("' + excelFilePath +'")\r';
// 	vbs += 'worksheetcount = objBook.Worksheets.Count\r';
// 	vbs += 'v = worksheetcount & ";"\r';
// 	// vbs += 's = CStr(v)\r';
// 	// vbs += 'msgbox s\r';
// 	vbs += 'objBook.close\r';
// 	vbs += 'Set objBook = Nothing\r';
//     vbs += 'Set objExcel = Nothing\r';
// 	vbs += 'Set objInDesign = CreateObject("InDesign.Application")\r';
// 	vbs += 'objInDesign.ScriptArgs.SetValue "excelData", v\r';
// 	vbs += 'End Function\r';
// 	vbs += 'ReadFromExcel()\r';

// 	app.doScript(vbs, ScriptLanguage.VISUAL_BASIC, undefined, UndoModes.FAST_ENTIRE_SCRIPT);
// 	var shtNumber = app.scriptArgs.getValue("excelData");
// 	app.scriptArgs.clear();
// 	return shtNumber
// }