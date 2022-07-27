#targetengine "session";
#include "functions/common.jsx";
#include "functions/packageBook.jsx";
#include "functions/packageDocument.jsx";

var main = new Window ("palette", "CE: 생활가전 Export Tool v.0.0.1");
var btn01 = main.add("button", [0, 0, 180, 30], "북 파일 패키지");
var btn03 = main.add("button", [0, 0, 180, 30], "북 파일 PDF 내보내기");
var btn02 = main.add("button", [0, 0, 180, 30], "문서 파일 패키지");
var btn04 = main.add("button", [0, 0, 180, 30], "문서 파일 PDF 내보내기");

var closeBook = "열려있는 북 파일을 닫고 실행하세요.";
var noBook = "열려있는 북 파일이 없습니다.";
var closeDocu = "열려있는 문서를 닫고 실행하세요.";
var noDocu = "열려있는 문서가 없습니다.";

btn01.onClick = function() {
	if (app.books.length == 0) {
		alert(noBook);
	} else if (app.documents.length > 0) {
		alert(closeDocu);
	} else {
		main.close();
		var Book = app.activeBook;
		preCheck(Book);
		var rs = packageBook(Book);
		if (rs == false) {
			alert("북 파일 패키지를 실패했습니다.");
		} else {
			alert("북 파일 패키지를 완료합니다.");
		}
		main.show();
	}
}

btn02.onClick = function() {
	if (app.documents.length == 0) {
		alert(noDocu);
	} else if (app.books.length > 0) {
		alert(closeBook);
	} else {
		main.close();
		var doc = app.activeDocument;
		var rs = packageDocument(doc);
		if (rs == false) {
			alert("문서 패키지를 실패했습니다.");
		} else {
			alert("문서 패키지를 완료합니다.");
		}
		main.show();
	}
}

btn03.onClick = function() {
	if (app.books.length == 0) {
		alert(noBook);
	} else {
		try {
			main.close();
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
			main.show();
		} catch(err) {
			alert(err.line + " : " + err);
			exit();
		}
	}
}

btn04.onClick = function() {
	if (app.documents.length == 0) {
		alert(noDocu);
	} else {
		try {
			main.close();
			var dVer = inputVersion();
			var cDate = GetDate();
			var selPreset = getData();
			var doc = app.activeDocument;
			var targetPath = doc.filePath;
			var targetName = doc.name.replace(/\.indd$/, "");
			var targetFile = new File(targetPath + "/" + targetName + cDate + "-" + dVer + ".pdf");
			doc.exportFile(ExportFormat.pdfType, targetFile, false, app.pdfExportPresets[selPreset]);
			alert ("완료합니다.");
			main.show();
		} catch(err) {
			alert(err.line + " : " + err);
			exit();
		}
	}
}

main.show();