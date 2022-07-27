/* Copyright 2019, Kasyan Servetsky
November 5, 2019,
Written by Kasyan Servetsky
http://www.kasyan.ho.com.ua
e-mail: askoldich@yahoo.com 
20200304 fneathe@gmail.com add export idml function */
//======================================================================================
// var scriptName = "인디자인 문서 패키지 + IDML",
// activeDocu;

// PreCheck();

//===================================== FUNCTIONS ======================================
function Main_packageDocu(activeDocu) {
	try { // if something goes wrong in the try-catch block, the batch processor won't stop here. It will log the error message and continue further
		var bookBasicName = activeDocu.name.replace(/\.indd$/, ""), // remove the extension from the document's name
		archiveFolderPath = activeDocu.filePath.absoluteURI, // the path to the folder where the active document is located
		packageFolderPath = archiveFolderPath + "/" + bookBasicName, // the path to the package folder
		packageFolder = new Folder(packageFolderPath); // the reference to the package folder
		if (!packageFolder.exists) packageFolder.create(); // create if it doesn't exist yet
			
		var result = activeDocu.packageForPrint(
						packageFolder, // to - File - The folder, alias, or path in which to place the packaged files.
						true, // copyingFonts - Boolean - If true, copies fonts used in the document to the package folder.
						true, // copyingLinkedGraphics - Boolean - If true, copies linked graphics files to the package folder. 
						false, // copyingProfiles - Boolean - If true, copies color profiles to the package folder. 
						true, // updatingGraphics - Boolean - If true, updates graphics links to the package folder. 
						true, // includingHiddenLayers - Boolean - If true, copies fonts and links from hidden layers to the package. 
						true, // ignorePreflightErrors - Boolean - If true, ignores preflight errors and proceeds with the packaging. If false, cancels the packaging when errors exist. 
						false, // creatingReport - Boolean - If true, creates a package report that includes printing instructions, print settings, lists of fonts, links and required inks, and other information. 
						//true, // includeIdml - Boolean - If true, generates and includes IDML in the package folder. (Optional)
						//true, // includePdf - Boolean - If true, generates and includes PDF in the package folder. (Optional)
						//"[High Quality Print]", // pdfStyle - String - If specified and PDF is to be included, use this style for PDF export if it is valid, otherwise use the last used PDF preset. (Optional)
						//false, // useDocumentHyphenationExceptionsOnly - Boolean - If this option is selected, InDesign flags this document so that it does not reflow when someone else opens or edits it on a computer that has different hyphenation and dictionary settings. (Optional)
						undefined, // versionComments - String - The comments for the version. (Optional) 
						undefined // forceSave - Boolean - If true, forcibly saves a version. (Optional) (default: false) 
					);
		if (result) {
			exportIDMLdocu(activeDocu)
			alert("인디자인 문서 패키지 + IDML 파일 생성을 완료합니다. ", scriptName, false);
		}
		else {
			alert("인디자인 문서 패키지를 실패했습니다.", scriptName, true);
		}
	}
	catch(err) {
		alert(err.message + ", line: " + err.line, scriptName, true);
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function PreCheckDocu() {
	if (app.documents.length == 0) alert_ErrorExit("열려있는 인디자인 문서가 없습니다.", true);
	var activeDocu = app.activeDocument;
	//CheckBook();
	Main_packageDocu(activeDocu);
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function alert_ErrorExit(error, icon) {
	alert(error, scriptName, icon);
	exit();
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function exportIDMLdocu(activeDocu) {
	var bookBasicName = activeDocu.name.replace(/\.indd$/, "");
	var myFolder = activeDocu.filePath.absoluteURI + "/" + bookBasicName;
	//app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
	//var myFolder = Folder.selectDialog( "인디자인 문서가 있는 상위 폴더를 선택하세요." );  
	var myFiles = [];
	GetSubFolders(myFolder);
	if ( myFiles.length > 0 ) {
		// turn off warnings: missing fonts, links, etc.
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		var myBookFileName = myFolder + "/" + myFolder.name;
		myBookFile = new File( myBookFileName );
			if ( myBookFile.exists ) {  
				if ( app.books.item(myFolder.displayName) == null ) {
						myBook = app.open( myBookFile );
				}
			}
			for ( i=0; i < myFiles.length; i++ ) {
				var myBookFile = app.open(File(myFiles[i]), true);
				with ( myBookFile ) {
					myFilePath = myFolder + "/" + myBookFile.name.split(".indd")[0] + ".idml"; //파일명 설정
					myFile = new File(myFilePath);
					myBookFile.exportFile(ExportFormat.INDESIGN_MARKUP, myFile); 
					myBookFile.close(SaveOptions.no)
				}
			}
	// turn on warnings
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;  
	//alert ("Complete");
	}

	//=================================== FUNCTIONS =========================================  
	function GetSubFolders(myFolder) {
		var my_folder = new Folder (myFolder)
		var myFileList = my_folder.getFiles();
		//var myFileList = [];
		for (var i = 0; i < myFileList.length; i++) {
			var myFile = myFileList[i];
			if (myFile instanceof Folder) {
				GetSubFolders(myFile);
			}
			else if (myFile instanceof File && myFile.name.match(/\.indd$/i)) {
				myFiles.push(myFile);
			}
		}
	}
}
