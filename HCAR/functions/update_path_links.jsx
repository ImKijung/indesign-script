/* Copyright 2014, Kasyan Servetsky
May 20, 2014
Written by Kasyan Servetsky
http://www.kasyan.ho.com.ua
e-mail: askoldich@yahoo.com

This script relinks old links with new ones in a folder you choose, including subfolders.
At first, it looks for the new link that has the same name as the old one, in the selected folder, and if it"s found, relinks it.
If the link with the exact name is not found and the link is not a text file (e.g. InCopy, Word, RTF, Exel file), 
the sript tries to find a file with the same base name but a different extension (".psd", ".tif", ".jpg", ".ai", ".eps", ".pdf", ".indd"). 
If the script doesn"t find the new link in the chosen folder, it repeats the above-mentioned process for every subfolder. */
//======================================================================================

var set, doc, links, linksFolder, subFolders, progressWin, progressBar, progressTxt, startTime,
scriptName = "인디자인 이미지 경로 재설정 유틸리티",
count = 0,
arr = [],
processedLinks = [],
extentions = [".psd", ".tif", ".jpg", ".ai", ".eps", ".pdf", ".indd"]; // List all possible file extensions



//===================================== FUNCTIONS  ======================================
function relinkMain() {
	var inddFolder, inddFiles;
	
	linksFolder = Folder.selectDialog("이미지가 들어있는 폴더를 선택하세요."); // global
	if (linksFolder == null) exit();
	subFolders = [linksFolder.fsName];
	GetSubFolders(linksFolder);

	CreateProgressBar();
	
	if (set.rbSel == 0 || set.rbSel == 1) {
		startTime = new Date();
		if (set.log) arr.push(GetDate());
		progressWin.show();
		if (set.rbSel == 0) { // active document
			doc = app.activeDocument;
			progressWin.show();
			ProcessDocument(doc, subFolders);
		}
		else if (set.rbSel == 1) { // all open documents
			for (var d = 0; d < app.documents.length; d++) {
				doc = app.documents[d];
				ProcessDocument(doc, subFolders);
				UpdateAllOutdatedLinks(doc);
			}
		}
	}
	else if (set.rbSel == 2) { // active book
		inddFiles = GetFilesFromBook();
		if (inddFiles.length == 0) ErrorExit("열려있는 북 파일에 인디자인 문서를 찾을 수 없습니다.", true);
		ProcessAllInddDocs(inddFiles);
	}
	else if (set.rbSel == 3 || set.rbSel == 4) {
		inddFolder = Folder.selectDialog("인디자인 문서가 들어있는 폴더를 선택하세요.");
		if (inddFolder == null) exit();
		progressWin.show();
		if (set.rbSel == 3) { // folder
			inddFiles = inddFolder.getFiles("*.indd");
		}
		else if (set.rbSel == 4) { // folder and subfolders
			inddFiles = GetAllInddFiles(inddFolder);
		}
		if (inddFiles.length == 0) ErrorExit("선택한 폴더에 인디자인 문서가 없습니다.", true);
		ProcessAllInddDocs(inddFiles);
	}
	
	progressWin.close();
	var endTime = new Date();
	var duration = GetDuration(startTime, endTime);
	var report = count + ((count == 1) ? " 개의 이미지가" : " 개의 이미지가") + " 재연결되었습니다.\n(경과 시간: " + duration + ")";

	if (set.log) {
		arr.push("\r=========================================================\r" + report + "\r=========================================================\r\r\r");
		var text = arr.join("\r");
		WriteToFile(text);
	}
	
	alert("완료했습니다. " + report, scriptName);
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function ProcessAllInddDocs(inddFiles) {
	startTime = new Date();
	if (set.log) {
		arr.push(GetDate());
		if (set.rbSel == 2) arr.push("북파일 이름: " + app.activeBook.name);
	}
	
	for (var i = 0; i < inddFiles.length; i++) {
		inddFile = inddFiles[i];
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
		doc = app.open(inddFile);
		app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
		ProcessDocument(doc, subFolders);
		UpdateAllOutdatedLinks(doc);
		doc.close(SaveOptions.YES);
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function ProcessDocument(doc, subFolders) {
	var oldLink, oldLinkFile, subFolderPath, currFolder, wasFound, basePath;
	if (set.log) arr.push("\r\r------------------------------------------\r문서 이름: " + doc.name + "\r문서 경로: " + File(doc.filePath).fsName + "\r------------------------------------------");
	
	links = doc.links; // global !!!
	var linksLength = links.length;
	
	if (linksLength == 0 || doc.name.match(/^Backup_/) != null) return; // This document doesn't contain any links
	if (set.backUp) {
		
		var file = new File(doc.filePath.fsName + "\\Backup_" + doc.name);
		if (file.exists) {
			var increment = 1;
			while (file.exists) {
				file = new File(doc.filePath.fsName + "\\Backup_" + increment++ + "_" + doc.name);
			}
		}
		doc.fullName.copy(file.fsName);
	}
	
	var pbCounter = 0;
	progressBar.minvalue = 0;
	progressBar.maxvalue = linksLength;

	for (var i = linksLength-1; i >= 0 ; i--) {
		oldLink = links[i];

		progressBar.value = pbCounter++;
		progressTxt.text = "파일 재 연결 중 - " + oldLink.name;
		
		if (LinkIsNotText(oldLink) && GetArrayIndex(processedLinks, oldLink.filePath) == -1) {
			oldLinkFile = new File(oldLink.filePath);
			if (set.onlyMissing && oldLink.status != LinkStatus.LINK_MISSING) continue;			
			wasFound = false;			
			if (set.log) arr.push("\r" + oldLink.name + "\r" + "\t이전 경로: " + oldLinkFile.fsName);
			
			for (var s = 0; s < subFolders.length; s++)  {
				subFolderPath = subFolders[s] + "/";
				currFolder = new Folder(subFolders[s]);
				
				progressWin.text = "폴더 안에 이미지들을 찾는 중 - \"" + currFolder.displayName + "\"";

				var newPath = currFolder.fsName + "/" + oldLink.name;
				
				var newFile =  new File(newPath);

				if (oldLinkFile.fsName != newPath) {
					if (newFile.exists) {
						Relink(oldLink, oldLinkFile, newFile);
						wasFound = true;
					}
					else { // If file not found, try with another extention
						for (var k = 0; k < extentions.length; k++) {
							if (newPath.match(/(.*)\.[^\.]+$/) != null) { // in case the link is without extension on Mac
								basePath = newPath.match(/(.*)\.[^\.]+$/)[1];
								newPath = basePath + extentions[k];
								newFile =  new File(newPath);
								//debugger
								if (newFile.exists) {
									Relink(oldLink, oldLinkFile, newFile);
									wasFound = true;
								}
                                }
						}
					}
				}
			} // subFolders.length
			if (!wasFound) arr.push("\t찾을 수가 없습니다.");		
		} // LinkIsNotText
	} // for (var i = links.length
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function Relink(oldLink, oldLinkFile, newFile) {
	var currentLink, currentLinkFile, allLinks,
	multiUsedLinks = [],	
	thisLinkUsedTimes = LinkUsage(oldLink);

	oldLink.relink(newFile);
	count++;
	processedLinks.push(newFile.fsName);
	if (set.log) arr.push("\t새 경로: " + newFile.fsName);

	try {
		oldLink = oldLink.update();
	}
	catch(err) {
		$.writeln(err.message + ", line: " + err.line);
	}

	if (thisLinkUsedTimes > 1) {
		allLinks = doc.links;
		if (set.log) arr.push("\tPlaced " + thisLinkUsedTimes + " times");

		for (var d = 0; d < allLinks.length; d++)  {
			currentLinkFile = new File(allLinks[d].filePath);
			if (currentLinkFile.fsName == oldLinkFile.fsName) {
				multiUsedLinks.push(allLinks[d].id);
			}
		}

		for (var f = multiUsedLinks.length-1; f >= 0 ; f--) {
			currentLink = doc.links.itemByID(multiUsedLinks[f]);
			if (currentLink != null) {
				currentLink.relink(newFile);
				count++;
			
				if (set.log && GetArrayIndex(arr, "\t새 경로: " + newFile.fsName) == -1) {
					arr.push("\t새 경로: " + newFile.fsName);
				}
				
				try {
					currentLink = currentLink.update();
				}
				catch(err) {
					$.writeln(err.message + ", line: " + err.line);
				}
			}
		}
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetAllInddFiles(folder) {
	var files = [],
	fileList = folder.getFiles(),
	i, file;
	
	for (i = 0; i < fileList.length; i++) {
		file = fileList[i];
		if (file instanceof Folder) {
			files = files.concat(GetAllInddFiles(file));
		}
		else if (file instanceof File && file.name.match(/\.indd$/i)) {
			files.push(file);
		}
	}

	return files;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetSubFolders(folder) {
	var file,
	fileList = folder.getFiles();
	
	for (var i = 0; i < fileList.length; i++) {
		file = fileList[i];
		
		if (file instanceof Folder) {
			subFolders.push(file.fsName);
			GetSubFolders(file);
		}
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function LinkUsage(link) {
	var counter = 0;
	
	for (var i =  0; i < links.length; i++) {
		if (link.filePath == links[i].filePath) {
			counter++;
		}
	}

	return counter;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function UpdateAllOutdatedLinks(doc) {
	var link;
	for (var i = doc.links.length-1; i >= 0; i--) {
		link = doc.links[i];
		if (link.status == LinkStatus.LINK_OUT_OF_DATE) link.update();
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function LinkIsNotText(link) {
	if (link.linkType != "InCopyInterchange" && 
		link.linkType != "InCopyStory" &&
		link.linkType != "InCopyMarkup" &&
		link.linkType != "Microsoft Excel Import Filter" &&
		link.linkType != "Microsoft Word Import Filter" &&
		link.linkType != "RTF Import Filter" &&
		link.linkType != "Text")
	{
		return true;
	}
	else {
		return false;
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetArrayIndex(arr, val) {
	for (var i = 0; i < arr.length; i++) {
		if (arr[i] == val) {
			return i;
		}
	}
	return -1;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetFilesFromBook() {
	var bookContent, file,
	activeBook = app.activeBook,
	files = [];
	
	for (var i = 0; i < activeBook.bookContents.length; i++) {
		bookContent = activeBook.bookContents[i];
		if (bookContent.status != BookContentStatus.MISSING_DOCUMENT) {
			file = new File(bookContent.fullName);
			files.push(file);
		}
	}
	
	return files;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function CreateDialog() {
	GetDialogSettings();
	var w = new Window("dialog", scriptName);
	
	w.p1 = w.add("panel", undefined, "Process");
	w.p1.orientation = "column";
	w.p1.alignment = "fill";
	w.p1.alignChildren = "left";
	
	//Radio buttons
	w.p1.rb0 = w.p1.add("radiobutton", undefined, "현재 열려있는 문서");
	if (app.documents.length == 0) w.p1.rb0.enabled = false;
	w.p1.rb1 = w.p1.add("radiobutton", undefined, "열려있는 모든 문서");
	if (app.documents.length < 2) w.p1.rb1.enabled = false;
	w.p1.rb2 = w.p1.add("radiobutton", undefined, "활성화 중인 Book 파일");
	if (app.books.length == 0) w.p1.rb2.enabled = false;
	w.p1.rb3 = w.p1.add("radiobutton", undefined, "인디자인 문서가 들어있는 폴더 선택");
	w.p1.rb4 = w.p1.add("radiobutton", undefined, "인디자인 문서가 들어있는 폴더들의 상위 폴더 선택");

	
	if (set.rbSel == 0 && app.documents.length != 0) {
		w.p1.rb0.value = true;
	}
	else if (set.rbSel == 1 && app.documents.length > 1) {
		w.p1.rb1.value = true;
	}
	else if (set.rbSel == 2 && app.books.length != 0) {
		w.p1.rb2.value = true;
	}
	else if (set.rbSel == 3) {
		w.p1.rb3.value = true;
	}
	else  {
		w.p1.rb4.value = true;
	}

	w.p2 = w.add("panel", undefined, "설정");
	w.p2.alignChildren = "left";
	w.p2.alignment = "fill";
	
	// Checkboxes
	w.p2.cb1 = w.p2.add("checkbox", undefined, "바탕화면에 로그 파일 만들기");
	w.p2.cb1.alignment = "left";
	w.p2.cb1.value = set.log;
	
	w.p2.cb2 = w.p2.add("checkbox", undefined, "잃어버린 이미지만 재 연결하기");
	w.p2.cb2.alignment = "left";
	w.p2.cb2.value = set.onlyMissing;
	
	w.p2.cb3 = w.p2.add("checkbox", undefined, "원본 인디자인 문서 백업하기");
	w.p2.cb3.alignment = "left";
	w.p2.cb3.value = set.backUp;
	
	// Buttons
	w.buttons = w.add("group");
	w.buttons.orientation = "row";   
	w.buttons.alignment = "center";
	w.buttons.ok = w.buttons.add("button", undefined, "OK", {name:"ok" });
	w.buttons.cancel = w.buttons.add("button", undefined, "Cancel", {name:"cancel"});
	
	var showDialog = w.show();
	
	if (showDialog == 1) {
		set.log = w.p2.cb1.value;
		set.onlyMissing = w.p2.cb2.value;
		set.backUp = w.p2.cb3.value;
		
		if (w.p1.rb0.value == true) {
			set.rbSel = 0;
		}
		else if (w.p1.rb1.value == true) {
			set.rbSel = 1;
		}
		else if (w.p1.rb2.value == true) {
			set.rbSel = 2;
		}
		else if (w.p1.rb3.value == true) {
			set.rbSel = 3;
		}
		else if (w.p1.rb4.value == true) {
			set.rbSel = 4;
		}
	
		app.insertLabel("Kas_" + scriptName, set.toSource());
		relinkMain();
	}
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetDialogSettings() {
	set = eval(app.extractLabel("Kas_" + scriptName));
	if (set == undefined) {
		set = { rbSel: 0, log: true, onlyMissing: false, backUp: true };
	}
	return set;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function CreateProgressBar() {
	progressWin = new Window("window", "이미지 파일을 재 연결하는 중..."); // global
	progressBar = progressWin.add("progressbar", [12, 12, 350, 24], 0, undefined); // global
	progressTxt = progressWin.add("statictext"); // global
	progressTxt.bounds = [0, 0, 340, 20];
	progressTxt.alignment = "left";
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetDuration(startTime, endTime) {
	var str;
	var duration = (endTime - startTime)/1000;
	duration = Math.round(duration);
	if (duration >= 60) {
		var minutes = Math.floor(duration/60);
		var seconds = duration - (minutes * 60);
		str = minutes + ((minutes != 1) ? " 분, " :  " 분, ") + seconds + ((seconds != 1) ? " 초" : " 초");
		if (minutes >= 60) {
			var hours = Math.floor(minutes/60);
			minutes = minutes - (hours * 60);
			str = hours + ((hours != 1) ? " 시간, " : " 시간, ") + minutes + ((minutes != 1) ? " 분, " :  " 분, ") + seconds + ((seconds != 1) ? " 초" : " 초");
		}
	}
	else {
		str = duration + ((duration != 1) ? " 초" : " 초");
	}

	return str;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function GetDate() {
	var date = new Date();
	if ((date.getYear() - 100) < 10) {
		var year = "0" + new String((date.getYear() - 100));
	}
	else {
		var year = new String((date.getYear() - 100));
	}
	var dateString = (date.getMonth() + 1) + "/" + date.getDate() + "/" + year + " " + date.getHours() + ":" + date.getMinutes() + ":" + date.getSeconds();
	return dateString;
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function WriteToFile(text) {
	var file = new File("~/Desktop/" + scriptName + ".txt");
	file.encoding = "UTF-8";
	if (file.exists) {
		file.open("e");
		file.seek(0, 2);
	}
	else {
		file.open("w");
	}
	file.write(text); 
	file.close();
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------
function ErrorExit(error, icon) {
	alert(error, scriptName, icon);
	exit();
}
//--------------------------------------------------------------------------------------------------------------------------------------------------------