/*Function to select the previous paragraph*/
function getPreviousPara(text){
    var nextPara = text.parent.characters.item(text.paragraphs[0].index-1).paragraphs[0];
    return nextPara
}

/*Global variables declaration*/
var myDoc = app.activeDocument;
var allChanges =myDoc.stories.everyItem().changes.everyItem().getElements();
var nChanges = allChanges.length;

/*Loop through all changes*/
for (var i = 0; i <= nChanges-1; i++) {
    /*Get story containing change*/
    var changeStory =  allChanges[i].parent;

    /*Apply character style to Added text and Deleted text*/
    if (allChanges[i].changeType == ChangeTypes.INSERTED_TEXT) {
        var cName = "Added";
        var mCstyle = app.activeDocument.characterStyles.item(cName);
        allChanges[i].characters.everyItem().applyCharacterStyle(mCstyle);
        allChanges[i].accept();
    } else if(allChanges[i].changeType == ChangeTypes.DELETED_TEXT) {
        var cName = "Deleted";
        var mCstyle = app.activeDocument.characterStyles.item(cName);
        var changeParent = allChanges[i].parent;
        allChanges[i].characters.everyItem().applyCharacterStyle(mCstyle);
        allChanges[i].reject();
    }

    /*Get the textStyleRanges of the story*/
    var everyStyleRange = changeStory.textStyleRanges;
    nItems = everyStyleRange.length;

    /*1st while loop variables declaration*/
    var found = 0;
    var i = 0;

    /*Loop through all textStyleRanges*/
    while(found==0 && i<=nItems-1){

        /*Look for a style range that has the Deleted or Added character style or the Title modif style*/
        if (everyStyleRange[i].appliedCharacterStyle.name == "Deleted" || everyStyleRange[i].appliedCharacterStyle.name == "Added" || everyStyleRange[i].appliedCharacterStyle.name == "Title modif"){

            /*When found, kill the loop*/
            found = 1;

            /*2nd while loop variables declaration*/
            var currentPara = everyStyleRange[i].paragraphs[0];
            var corrected = 0;
            var maxLoopLimit = 0;

            /*Loop through all paragraphs of the story going backward from the found style range*/
            while(corrected==0 && maxLoopLimit < changeStory.paragraphs.length){
                maxLoopLimit++;

                /*Get the paragraph style name*/
                var currentParaStyleName= currentPara.appliedParagraphStyle.name;

                /*Look for the Title paragraph style applied*/
                if(currentParaStyleName == "Title"){

                    /*Title found, apply new paragraph style and kill 2nd while loop*/
                    currentPara.appliedParagraphStyle = currentParaStyleName+" modif";
                    corrected++; 
                }
                else if(currentParaStyleName == "Title modif"){

                    /*Title already changed, kill 2nd while loop*/
                    corrected++;
                }
                else{

                    /*Get previous paragraph and start loop again, acts as counter*/
                     currentPara = getPreviousPara(currentPara);
                }

            /*2nd while loop end*/
            }

        /*if end*/
        }

        /*If no modification were found in the story, alert the user that an error occured*/
        if(i>=nItems){alert("No changes character style was found so no Title change occured")}

        /*increment 1st loop counter*/
        i++;

    /*1st loop end*/
    }
}