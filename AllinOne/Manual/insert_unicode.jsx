function insertLtoR() {
	if (app.selection[0].constructor.name == "InsertionPoint"){ 
		with (app.selection[0].insertionPoints[0]){ // insertion point 
			contents = String.fromCharCode(0x200E); // dotless i 
		} 
	}
}

function insertRtoL() {
	if (app.selection[0].constructor.name == "InsertionPoint"){ 
		with (app.selection[0].insertionPoints[0]){ // insertion point 
			contents = String.fromCharCode(0x200F); // dotless i 
		} 
	}
}