
var doc = app.activeDocument;

//이미지에 Tag 적용하기
var myStory = doc.stories;
for (var y=0; y < myStory.length; y++) {
     var myGraphics = myStory[y].allGraphics;
     for (x=0; x < myGraphics.length; x++){
          myGraphics[x].autoTag();
     }
}