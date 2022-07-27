#include "common.jsx";

if (app.documents.length == 0) {
	alert(noDocu);
} else if (app.books.length > 0) {
	alert(closeBook);
} else {
	var doc = app.activeDocument;
	var rs = packageDocument(doc);
	if (rs == false) {
		alert("문서 패키지를 실패했습니다.");
	} else {
		alert("문서 패키지를 완료합니다.");
	}
}

function packageDocument(doc) {
	try {
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var dVer = inputVersion();
		var cDate = GetDate();
		var appVer = GetAppVersion();
		var docName = doc.name.replace(/\.indd$/, "");
		var docPath = doc.filePath.absoluteURI;
		var packagePath = docPath + "/" + docName + dVer + "_" + cDate;
		var packageFolder = new Folder(packagePath);

		if (!packageFolder.exists) packageFolder.create();

		var result = doc.packageForPrint(
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
		)
		if (result) {
			var wReport = new File(packagePath + "/" + appVer);
			wReport.open('w');
			wReport.close();
			var oDocFile = new File (packagePath + "/" + doc.name);
			var oDoc = app.open(oDocFile);
			eXportIDML(packagePath, oDoc);
			oDoc.close(SaveOptions.NO);
			app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
			return true
		} else {
			alert("문서 패키지를 실패했습니다.");
			return false
		}

	} catch(err) {
		alert(err.line + ":" + err);
		exit();
	}
}