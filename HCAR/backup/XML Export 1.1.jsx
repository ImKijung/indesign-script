// Copyright 2011 L'Express de Toronto Inc.
// December 4, 2011
// Written by Kasyan Servetsky
// http://www.kasyan.ho.com.ua
// e-mail: askoldich@yahoo.com
//======================================================================================
const gScriptName = "XML Export";
const gScriptVersion = "1.1";

if (app.documents.length == 0) ErrorExit("Please open a document and try again.");

var s = GetSettings();

CreateDialog();

//===================================== FUNCTIONS  ======================================
function Main() {
	var doc = app.activeDocument;
	var issueFolder = doc.filePath.parent;
	var webFolderPath = issueFolder.absoluteURI + "/web/";
	var webFolder = new Folder(webFolderPath);
	if (!webFolder.exists) ErrorExit("Folder \"" + File.decode(webFolder.name) + "\" doesn't exist.");
	
	var xmlFilePath = webFolderPath + doc.name.replace(/\.indd$/i, "") + ".xml"; 
	var xmlFile = new File(xmlFilePath);
	
	if (s.mapStylesToTags) {
		MapStylesToTags(doc);
	}

	with(doc.xmlExportPreferences) {
		// General
		exportUntaggedTablesFormat = XMLExportUntaggedTablesFormat.CALS;
		characterReferences = false; // Remap Break, Whitespace, and Special Characters		
		fileEncoding = XMLFileEncoding.UTF8;
		// Images
		copyOriginalImages = false;
		copyOptimizedImages = false;
		copyFormattedImages = false;
	}

	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.neverInteract;
	doc.exportFile(ExportFormat.XML, xmlFile, false);
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;	
}
//---------------------------------------------------------------------------------------------------------------------------------------------------------
function MapStylesToTags(doc) {
	var parStyles = doc.allParagraphStyles;
	var charStyles = doc.allCharacterStyles;
	var xmlTags =  doc.xmlTags;

	for (var i = 0; i < parStyles.length; i++) {
		for (var t = 0; t < xmlTags.length; t++) {
			if (parStyles[i].name == xmlTags[t].name) {
				doc.xmlExportMaps.add(parStyles[i], xmlTags[t]);
			}
		}
	}

	for (var c = 0; c < charStyles.length; c++) {
		for (var x = 0; x < xmlTags.length; x++) {
			if (charStyles[c].name == xmlTags[x].name) {
				doc.xmlExportMaps.add(charStyles[c], xmlTags[x]);
			}
		}
	}

	doc.mapStylesToXMLTags();
}
//---------------------------------------------------------------------------------------------------------------------------------------------------------
function CreateDialog() {
	var w = new Window("dialog", gScriptName + " - " + gScriptVersion);
	w.p = w.add("panel", undefined, undefined);
	// Checkbox
	w.p.cb = w.p.add("checkbox", undefined, "Map Styles to Tags");
	w.p.cb.value = s.mapStylesToTags;
	// Buttons
	w.buttons = w.add("group");
	w.buttons.orientation = "row";
	w.buttons.alignment = "center";
	w.buttons.ok = w.buttons.add("button", undefined, localize({en:"OK", fr:"OK"}), {name:"ok" });
	w.buttons.cancel = w.buttons.add("button", undefined, localize({en:"Cancel", fr:"Annuler"}), {name:"cancel"});

	var showDialog = w.show();
	
	if (showDialog == 1) {
		s.mapStylesToTags = w.p.cb.value;
		app.insertLabel("Kas_" + gScriptName + gScriptVersion, s.toSource());
		Main();
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetSettings() {
	var set = eval(app.extractLabel("Kas_" + gScriptName + gScriptVersion));
	if (set == undefined) {
		//$.writeln("Creating default settings");
		set = { mapStylesToTags : true };
	}
	
	return set;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function ErrorExit(error, icon) {
	alert(error, gScriptName + " - " + gScriptVersion, icon);
	exit();
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------