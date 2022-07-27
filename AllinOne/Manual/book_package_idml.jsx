// Package for Archive.jsx
// Script for InDesign CS3, CS4 and CS5 -- packages all InDesign documents in the selected folder.
// Version 4.0
// April 1 2011
// Written by Kasyan Servetsky
// http://www.kasyan.ho.com.ua
// e-mail: askoldich@yahoo.com
//---------------------------------------------------------------------------------------- ----------------------
const gScriptName = "PackageForArchive";
const gScriptVersion = "4.0";
 
 
var myInDesignVersion = Number(String(app.version).split(".")[0]);
 
 
var myFolder = Folder.selectDialog("Select a folder for package");
if (myFolder == null) exit();
var myFilelist = [];
var myAllFilesList = myFolder.getFiles();
 
 
for (var f = 0; f < myAllFilesList.length; f++) {
  var myFile = myAllFilesList[f];
  if (myFile instanceof File && myFile.name.match(/\.indb$/i)) {
  myFilelist.push(myFile);
  }
}
 
 
if (myFilelist.length == 0) {
  alert("No files to open.", "Package for archive script");
  exit();
}
 
 
var myDialogResult = CreateDialog();
if (myDialogResult == undefined) {
  exit();
}
 
 
if (myDialogResult.createLogCheckBox) {
  WriteToFile("\r--------------------- Script started -- " + GetDate() + " ---------------------\n");
  WriteToFile("\rSelected folder -- " + myFolder.fsName.replace("/Volumes/", "") + "\n");
}
 
 
var myOutFolder = new Folder( myFolder.fsName + "/Archive/" );
VerifyFolder(myOutFolder);
 
 
var myBackUpFolder = new Folder( myFolder.fsName + "/BackUp/" );
VerifyFolder(myBackUpFolder);
 
 
if (myFolder.fsName == myOutFolder.fsName) exit();
 
 
var myCounter = 1;
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
 
 
  // Progress bar
  if (myInDesignVersion >= 5) {
  var myProgressWin = new Window ( 'window', 'Packaging files' );
  var myProgressBar = myProgressWin.add ('progressbar', [12, 12, 350, 24], 0, myFilelist.length);
  var myProgressTxt = myProgressWin.add('statictext', undefined, 'Starting packaging files');
  myProgressTxt.bounds = [0, 0, 340, 20];
  myProgressTxt.alignment = "left";
  myProgressWin.show();
  } // Progress bar
 
 
for (var i = myFilelist.length-1; i >= 0; i--) {
  var myCurrentFile = myFilelist[i];
  var myNewName = GetNameWithoutExtension(myCurrentFile);
 
  try {
  var myDoc = app.open(myCurrentFile, false);
  var myDocName = myDoc.name;
 
 
  // Progress bar
  if (myInDesignVersion >= 5) {
  myProgressBar.value = myCounter;
  myProgressTxt.text = String("Packaging file - " + myDocName + " (" + myCounter + " of " + myFilelist.length + ")");
  } // Progress bar
 
  UpdateAllOutdatedLinks();
 
  var myNewFolder = new Folder(myOutFolder.fsName + "/" + myNewName);
  VerifyFolder(myNewFolder);
 
  if (myInDesignVersion == 5) {
  var myPackageOk = myDoc.packageForPrint(myNewFolder, myDialogResult.copyFontsCheckBox, myDialogResult.copyGraphicsCheckBox, false, myDialogResult.updateGraphicsCheckBox, myDialogResult.ignorePreflightErrorsCheckBox, myDialogResult.createReportCheckBox);
  }
  else if (myInDesignVersion > 5) {
  var myPackageOk = myDoc.packageForPrint(myNewFolder, myDialogResult.copyFontsCheckBox, myDialogResult.copyGraphicsCheckBox, false, myDialogResult.updateGraphicsCheckBox, myDialogResult.includeHiddenLayers, myDialogResult.ignorePreflightErrorsCheckBox, myDialogResult.createReportCheckBox);
  }
 
  if (myPackageOk && myDialogResult.createLogCheckBox) {
  WriteToFile(myCounter + " - " + myCurrentFile.fsName.replace("/Volumes/", "") + " -- Ok\n");
  myCounter++;
  CatchMissingLinks(myDoc);
  }
 
  myDoc.close(SaveOptions.NO);
  var myMoved = MoveFile(myCurrentFile, myBackUpFolder);
  if (!myMoved && myDialogResult.createLogCheckBox) WriteToFile(myCounter + " - ERROR  -- Can't move \"" + myCurrentFile.fsName.replace("/Volumes/", "") + "\" to BackUp folder\n");
 
 
  }
  catch(e) {
  if (myDialogResult.createLogCheckBox) WriteToFile(myCounter + " - ERROR  -- " + myCurrentFile.fsName.replace("/Volumes/", "") + " - " + e + "\n");
  myCounter++;
  continue;
  }
 
 
  } // end For loop
 
 
  // Progress bar
  if (myInDesignVersion >= 5) {
  myProgressWin.close();
  } // Progress bar
 
 
app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
if (myDialogResult.createLogCheckBox) WriteToFile("\r--------------------- Script finished -- " + GetDate() + " ---------------------\r\r");
 
 
alert("Done.", "Package for archive script");
 
 
// ------------------------------------------------- FUNCTIONS -------------------------------------------------
function UpdateAllOutdatedLinks() {
  for (var j = myDoc.links.length-1; j >= 0; j--) {
  var myLink = myDoc.links[j];
  if (myLink.status == LinkStatus.linkOutOfDate){
  try {
  myLink.update();
  }
  catch(e) {
  if (myDialogResult.createLogCheckBox) WriteToFile("\tCAN'T UPDATE LINK  -- " + myLink.name + "\n");
  }
  }
  }
}
//---------------------------------------------------------------------------------------- ----------------------
function WriteToFile(myText) {
  myFile = new File("~/Desktop/Package for Archive Report.txt");
  if ( myFile.exists ) {
  myFile.open("e");
  myFile.seek(0, 2);
  }
  else {
  myFile.open("w");
  }
  myFile.write(myText);
  myFile.close();
}
//---------------------------------------------------------------------------------------- ----------------------
function GetDate() {
  var myDate = new Date();
  if ((myDate.getYear() - 100) < 10) {
  var myYear = "0" + new String((myDate.getYear() - 100));
  } else {
  var myYear = new String ((myDate.getYear() - 100));
  }
  var myDateString = (myDate.getMonth() + 1) + "/" + myDate.getDate() + "/" + myYear + " " + myDate.getHours() + ":" + myDate.getMinutes() + ":" + myDate.getSeconds();
  return myDateString;
}
//---------------------------------------------------------------------------------------- ----------------------
// GetNameWithoutExtension and VerifyFolder functions were written on basis of Bob Stucky's
// extensions to object prototypes found in his Script Library
function GetNameWithoutExtension(myFile) {
  var myFileName = myFile.name;
  var myIndex = myFileName.lastIndexOf( "." );
  if ( myIndex > -1 ) {
  myFileName = myFileName.substr(0, myIndex);
  }
  return myFileName;
}
//---------------------------------------------------------------------------------------- ----------------------
function VerifyFolder(myFolder) {
  if (!myFolder.exists) {
  var myFolder = new Folder(myFolder.absoluteURI);
  var myArray1 = new Array();
  while (!myFolder.exists) {
  myArray1.push(myFolder);
  myFolder = new Folder(myFolder.path);
  }
  myArray2 = new Array();
  while (myArray1.length > 0) {
  myFolder = myArray1.pop();
  if (myFolder.create()) {
  myArray2.push(myFolder);
  } else {
  while (myArray2.length > 0) {
  myArray2.pop.remove();
  }
  throw "Folder creation failed";
  }
  }
  }
}
//---------------------------------------------------------------------------------------- ----------------------
function CatchMissingLinks(myDoc) {
  var myMissingLinks = new Array;
  var myLinks = myDoc.links;
  var myHeader = false;
  for (var i = 0; i < myLinks.length; i++) {
  if (myLinks[i].status == LinkStatus.linkMissing) {
  var myLinkName = myLinks[i].name;
  if (!myHeader) {
  WriteToFile("\tMISSING LINKS:\r");
  myHeader = true;
  WriteToFile("\t" + myLinkName + "\r");
  }
  }
  }
}
//---------------------------------------------------------------------------------------- ----------------------
function CreateDialog() {
  var mySettings = GetSettings();
  var myDialog = new Window('dialog', 'Package for Archive');
  myDialog.orientation = 'column';
  myDialog.alignChildren = 'top';
 
  var myFolderPanel = myDialog.add('panel', undefined, 'Folder to package:');
  myFolderPanel.alignment = 'fill';
  var myFolderPanelStTxt1 = myFolderPanel.add('statictext', undefined, myFolder.fsName.replace("/Volumes/", ""));
  var myFolderPanelStTxt2 = myFolderPanel.add('statictext', undefined, '(Found ' + myFilelist.length + ' files)');
 
  var myCheckBoxPnl = myDialog.add('panel', undefined, 'Options');
  myCheckBoxPnl.orientation = 'row';
 
  var myCheckBoxGroup1 = myCheckBoxPnl.add('group');
  myCheckBoxGroup1.orientation = 'column';
  myCheckBoxGroup1.alignChildren = 'left';
  var myCheckBoxGroup2 = myCheckBoxPnl.add('group');
  myCheckBoxGroup2.orientation = 'column';
  myCheckBoxGroup2.alignChildren = 'left';
 
  var myCopyFontsCheckBox = myCheckBoxGroup1.add('checkbox', undefined, 'Copy fonts');
  myCopyFontsCheckBox.value = mySettings.copyFontsCheckBox;
  myCopyFontsCheckBox.helpTip = "If checked, copies fonts used in the document to the package folder.";
 
  var myCopyGraphicsCheckBox = myCheckBoxGroup1.add('checkbox', undefined, 'Copy linked graphics');
  myCopyGraphicsCheckBox.value = mySettings.copyGraphicsCheckBox;
  myCopyGraphicsCheckBox.helpTip = "If checked, copies linked graphics files to the package folder.";
 
  var myUpdateGraphicsCheckBox = myCheckBoxGroup1.add('checkbox', undefined, 'Update graphics');
  myUpdateGraphicsCheckBox.value = mySettings.updateGraphicsCheckBox;
  myUpdateGraphicsCheckBox.helpTip = "If checked, updates graphics links to the package folder.";
 
  if (myInDesignVersion > 5) {
  var myIncludeHiddenLayersCheckBox = myCheckBoxGroup1.add('checkbox', undefined, 'Include hidden layers');
  myIncludeHiddenLayersCheckBox.value = mySettings.includeHiddenLayers;
  myIncludeHiddenLayersCheckBox.helpTip = "If checked, copies fonts and links from hidden layers to the package.";
  }
 
  var myIgnorePreflightErrorsCheckBox = myCheckBoxGroup2.add('checkbox', undefined, 'Ignore preflight errors');
  myIgnorePreflightErrorsCheckBox.value = mySettings.ignorePreflightErrorsCheckBox;
  myIgnorePreflightErrorsCheckBox.helpTip = "If checked, ignores preflight errors and proceeds with the packaging. If not, cancels the packaging when errors exist.";
 
  var myCreateReportCheckBox = myCheckBoxGroup2.add('checkbox', undefined, 'Create report');
  myCreateReportCheckBox.value = mySettings.createReportCheckBox;
  myCreateReportCheckBox.helpTip = "If checked, creates a package report that includes printing instructions, print settings, lists of fonts, links and required inks, and other information.";
 
  var myCreateLogCheckBox = myCheckBoxGroup2.add('checkbox', undefined, 'Create log file on the desktop');
  myCreateLogCheckBox.value = mySettings.createLogCheckBox;
  myCreateLogCheckBox.helpTip = "If checked, creates a log file on the desktop that includes error messages";
 
  var myOkCancelGroup = myDialog.add('group');
  myOkCancelGroup.orientation = 'row';
  myOkCancelGroup.alignment = 'center';
  var myOkBtn = myOkCancelGroup.add('button', undefined, 'Go', {name:'ok'});
  var myCancelBtn = myOkCancelGroup.add('button', undefined, 'Quit', {name:'cancel'});
 
  var myShowDialog = myDialog.show();
 
  if (myShowDialog== 1) {
  var myResult = {};
  myResult.copyFontsCheckBox = myCopyFontsCheckBox.value;
  myResult.copyGraphicsCheckBox = myCopyGraphicsCheckBox.value;
  myResult.updateGraphicsCheckBox = myUpdateGraphicsCheckBox.value;
  if (myInDesignVersion > 5) myResult.includeHiddenLayers = myIncludeHiddenLayersCheckBox.value;
  myResult.ignorePreflightErrorsCheckBox = myIgnorePreflightErrorsCheckBox.value;
  myResult.createReportCheckBox = myCreateReportCheckBox.value;
  myResult.createLogCheckBox = myCreateLogCheckBox.value;
 
  myResult.folder = myFolder;
  app.insertLabel("Kas_" + gScriptName + gScriptVersion, myResult.toSource());
  }
  return myResult;
}
//---------------------------------------------------------------------------------------- ----------------------
function GetSettings() {
  var mySettings = eval(app.extractLabel("Kas_" + gScriptName + gScriptVersion));
 
  if (mySettings == undefined) {
  mySettings = { copyFontsCheckBox:false, copyGraphicsCheckBox:true, updateGraphicsCheckBox:true, includeHiddenLayers:true, ignorePreflightErrorsCheckBox:true, createReportCheckBox:false, createLogCheckBox:true };
  }
  return mySettings;
}
//---------------------------------------------------------------------------------------- ----------------------
function MoveFile(myFile, myFolder) {
  if (!myFile instanceof File || !myFolder instanceof Folder || !myFile.exists || !myFolder.exists) return false;
  var myMovedFile = new File(myFolder.absoluteURI + "/" + myFile.name);
  if (File.fs == "Windows")  {
  var myVbScript = 'Set fs = CreateObject("Scripting.FileSystemObject")\r';
  myVbScript +=  'fs.MoveFile "' + myFile.fsName + '", "' + myFolder.fsName + '\\"';
  app.doScript(myVbScript, ScriptLanguage.visualBasic);
  }
  else if (File.fs == "Macintosh") {
  if (myFolder.fsName.match("Desktop") != null) {
  var myMacFolder = myFolder.fsName.replace(/\//g, ":");
  var myMacFile = myFile.fsName.replace(/\//g, ":");
  var myAppleScript = 'tell application "Finder"\r';
  myAppleScript += 'set myFolder to a reference to folder (name of startup disk & "' + myMacFolder + '")\r';
  myAppleScript += 'set myFile to a reference to file (name of startup disk & "' + myMacFile + '")\r';
  myAppleScript += 'tell myFile\r';
  myAppleScript += 'move to myFolder\r';
  myAppleScript += 'end tell\r';
  myAppleScript += 'end tell\r';
  }
  else if (myFolder.fsName.match("Documents") != null) {
  var myMacFolder = myFolder.fsName.replace(/\//g, ":");
  var myMacFile = myFile.fsName.replace(/\//g, ":");
  var myAppleScript = 'tell application "Finder"\r';
  myAppleScript += 'set myFolder to a reference to folder (name of startup disk & "' + myMacFolder + '")\r';
  myAppleScript += 'set myFile to a reference to file (name of startup disk & "' + myMacFile + '")\r';
  myAppleScript += 'tell myFile\r';
  myAppleScript += 'move to myFolder\r';
  myAppleScript += 'end tell\r';
  myAppleScript += 'end tell\r';
  }
  else {
  var myMacFolder = myFolder.fullName.replace(/\//, "").replace(/\//g, ":");
  var myMacFile = myFile.fullName.replace(/\//, "").replace(/\//g, ":");
  var myAppleScript = 'tell application "Finder"\r';
  myAppleScript += 'set myFolder to folder "' + myMacFolder + '"\r';
  myAppleScript += 'set myFile to document file "' + myMacFile + '"\r';
  myAppleScript += 'tell myFile\r';
  myAppleScript += 'move to myFolder\r';
  myAppleScript += 'end tell\r';
  myAppleScript += 'end tell\r';
  }
  app.doScript(myAppleScript, ScriptLanguage.applescriptLanguage);
  }
  if (myMovedFile.exists) {
  return true;
  }
  else {
  return false;
  }
}