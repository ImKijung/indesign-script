var myFolder = Folder.selectDialog( "Select a folder with InDesign files" );
if ( myFolder != null ) {
	var myFiles = [];
	var myAllFilesList = myFolder.getFiles();

	for (var f = 0; f < myAllFilesList.length; f++) {
		var myFile = myAllFilesList[f];
		if (myFile instanceof File && myFile.name.match(/\.indd$/i)) {
			myFiles.push(myFile);
		}
	}
	
	if ( myFiles.length > 0 ) {
		var myBookFileName = myFolder + "/"+ myFolder.name + ".indb";
		myBookFile = new File( myBookFileName );
		if ( myBookFile.exists ) {
			if ( app.books.item(myFolder.displayName + ".indb") == null ) {
				myBook = app.open( myBookFile );
			}
		}
		else {
			 myBook = app.books.add( myBookFile );
			 myBook.automaticPagination = false;
			 for ( i=0; i < myFiles.length; i++ ) {
				myBook.bookContents.add( myFiles[i] );
			 }
			 myBook.save( );
	 	}
	}
}