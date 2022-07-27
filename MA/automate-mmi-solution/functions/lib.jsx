function readTabDelimitedFile(fPath) {
	var returnArray = new Array();
	//-- Verify that the file exists
	var fileObject = File(fPath);
	if (!fileObject.exists) {
		return returnArray ; // an empty array because the file doesn't exist.
	}
	//-- Create a regular expression for a tab.
	var tabExpression = new RegExp('\\t') ;
	//-- Read the file.
	try {
		//-- The file has to be open.
		fileObject.open('r'); //-- Open for reading.
		while (!fileObject.eof) {
			var currentLine = fileObject.readln();
			if (tabExpression.test(currentLine) ) {
				returnArray.push(currentLine.split('\t')) ;
			}
		}
		fileObject.close();
	}
	catch(errMain) {
		try {
			fileObject.close() ;
		}
		//-- if the close generates an error skip it.
		catch(errInner) { /* nothing here */ }
	}
	return returnArray ;
	var group = returnArray;
}