#include "common.jsx";

if (app.books.length == 0) {
	alert(noBook);
} else {
	try {
		openDocuments();
		var dVer = inputVersion();
		var cDate = GetDate();
		var selPreset = getData();
		var Book = app.activeBook;
		var targetPath = Book.filePath;
		var targetName = Book.name.replace(/\.indb$/, "");
		var targetFile = new File(targetPath + "/" + targetName + cDate + "-" + dVer + ".pdf");
		Book.exportFile(ExportFormat.pdfType, targetFile, false, app.pdfExportPresets[selPreset]);
		closeDocs();
		alert ("완료합니다.");
	} catch(err) {
		alert(err.line + " : " + err);
		exit();
	}
}