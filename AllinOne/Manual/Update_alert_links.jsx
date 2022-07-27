function Update_alert_links() {
     if (app.documents.length == 0) {  
          alert("인디자인 문서를 열고 실행해주세요.");  
          exit();  
     }  
          
     var docs = app.documents;  
          
     for (var i = 0; i < docs.length; i++) {  
          UpdateAllOutdatedLinks(docs[i]);  
     }  
          
     alert("완료합니다.");  
          
     function UpdateAllOutdatedLinks(doc) {  
          for (var d = doc.links.length-1; d >= 0; d--) {  
               var link = doc.links[d];  
               if (link.status == LinkStatus.linkOutOfDate) {  
                    link.update();  
               }  
          }  
     }  
}