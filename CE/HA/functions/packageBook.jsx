#include "common.jsx";

if (app.books.length == 0) {
	alert(noBook);
} else if (app.documents.length > 0) {
	alert(closeDocu);
} else {
	var sBook = app.activeBook;
	preCheck(sBook);
	var rs = packageBook(sBook);
	if (rs == false) {
		alert("북 파일 패키지를 실패했습니다.");
	} else {
		alert("북 파일 패키지를 완료합니다.");
	}
}

function packageBook(Book) {
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	try {
		var dVer = inputVersion();
		var cDate = GetDate();
		var appVer = GetAppVersion();
		var bookName = Book.name.replace(/\.indb$/, "");
		var bookPath = Book.filePath.absoluteURI;
		var packagePath = bookPath + "/" + bookName + dVer + "_" + cDate;
		var packageFolder = new Folder(packagePath);

		if (!packageFolder.exists) packageFolder.create();
		var result = Book.packageForPrint(
			packageFolder, // package folder
			true, // copy fonts
			true, // copy links
			false, // copy profiles
			true, // update graphics
			true, // include hidden layers
			true, // ignorePreflightError
			false, // create Report
			//true, // includeIdml
			//true, // includePdf
			//"[High Quality Print]" //pdfStyle
			//false, //useDocumentHyphenationExceptionsOnly
			//undefined // vserion Comments
			//undefined // forceSave
		);

		if (result) {
			var wReport = new File(packagePath + "/" + appVer);
			wReport.open('w');
			wReport.close();
			var cDoc;
			var myWindow = new Window ('palette');
				myWindow.pbar = myWindow.add ('progressbar', undefined, 0, Book.bookContents.length);
				myWindow.pbar.preferredSize.width = 300;
				myWindow.show();
			
			for (var j=0; j<Book.bookContents.length; j++) {
				myWindow.pbar.value = j + 1;
				cBookContent = Book.bookContents[j];
				cBookFile = new File(packagePath + "/" + cBookContent.name);
				cDoc = app.open(cBookFile);
				eXportIDML(packagePath, cDoc);
				cDoc.close(SaveOptions.NO);
			}
			app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
			return true
		} else {
			return false
		}
	} catch(err) {
		alert(err.line + ":" + err);
		exit();
	}
}