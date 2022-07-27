#include "common.jsx";

if (app.documents.length == 0) {
	alert(noDocu);
} else {
	try {
		var dVer = inputVersion();
		var cDate = GetDate();
		var selPreset = getData();
		var doc = app.activeDocument;
		var targetPath = doc.filePath;
		var targetName = doc.name.replace(/\.indd$/, "");
		var targetFile = new File(targetPath + "/" + targetName + cDate + "-" + dVer + ".pdf");
		doc.exportFile(ExportFormat.pdfType, targetFile, false, app.pdfExportPresets[selPreset]);
		alert ("완료합니다.");
	} catch(err) {
		alert(err.line + " : " + err);
		exit();
	}
}