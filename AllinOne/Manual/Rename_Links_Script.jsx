// Rename Link Scritp 최종본
// 속도 개선
function renameLinksScript() {
    if(app.documents.length < 1){
        alert("인디자인 문서를 여세요!");
        exit();
    } else {
        sourceFolder = Folder.selectDialog("이미지 파일이 있는 폴더 경로를 선택하세요.\r폴더를 잘못 선택했을 경우 오류가 날 수 있습니다.")
        if (sourceFolder != null) {
            renameLinkMain (sourceFolder);
            alert ("파일명 변환 완료");
        } else {
            exit();
        }
    }
}

function renameLinkMain(sourceFolder) {
    var myDoc = app.activeDocument;
    var myLinks = myDoc.links;
    for (i = myLinks.length-1; i >= 0 ; i--) {
        var myLink = myLinks[i];
        var myFile = new File(myLink.filePath);
        
        var linkName = myLink.name;
        var oldLink = myLink;
        var oldLinkName = linkName;
        var ext = linkName.substr(linkName.lastIndexOf("."));
        var newName1 = linkName.split("-").join("_");
        var newName2 = newName1.split(" ").join("_");
        var newName3 = newName2.split("+").join("_");
        var newName4 = newName3.split("(").join("_");
        var newName5 = newName4.split(").").join(".");
        var newName6 = newName5.split(")").join("_");
        var newName7 = newName6.split("_.").join(".");
        var imgName = newName7.substr(0, newName7.lastIndexOf("."));
        var newName8 = imgName.toString().split(",").join("_");
        var newName = newName8.toString().split(".").join("_") + ext;
        var newFile = sourceFolder + "/" + newName;
        if (myLink.status == LinkStatus.NORMAL){
            myFile.rename(newName);
            myLink.relink(myFile);
        }
        else if (myLink.status == LinkStatus.LINK_MISSING){
            imglist = findAnyFileDownThere(newName, sourceFolder);
            if (imglist.length == 1) {
                myLink.relink(imglist[0]);
            }
        } 
    }
}

function findAnyFileDownThere (filename, path) {  
    var result, folderList, fl;  
    result = Folder(path).getFiles (filename+".*");  
    if (result.length > 0)  
    return result;  
    folderList = Folder(path).getFiles(function(item) { return item instanceof Folder && !item.hidden; });  
    
    for (fl=0; fl<folderList.length; fl++) {  
        $.writeln("Looking in: "+path+"/"+folderList[fl].name);  
        result = findAnyFileDownThere (filename, path+"/"+folderList[fl].name);  
        if (result.length > 0)  
        return result;  
    }  
    return [];  
} 