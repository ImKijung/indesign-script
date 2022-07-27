function RemoveBlankPages () {
	var doc = app.documents;
	for (k=0; k < doc.length; k++) {
		var pages = doc[k].pages.everyItem().getElements();
		for(var i = pages.length-1;i>=0;i--){
			var removePage = true;
			if(pages[i].pageItems.length>0){
				var items = pages[i].pageItems.everyItem().getElements();
				for(var j=0;j<items.length;j++){
					if(!(items[j] instanceof TextFrame)){removePage=false;break}
					if(items[j].contents!=""){removePage=false;break}
				}
			}
			if(i==0 && doc[k].pages.length==1){
				removePage = false;
				}
			if(removePage){
				pages[i].remove();
				//exit();
				}
		}
	} alert ("빈 페이지를 삭제했습니다.");
}