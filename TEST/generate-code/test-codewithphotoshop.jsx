#targetengine "session"

var outputfile = "c:/Users/adfasdfasdf/Desktop/GH789845.png"
runPhotoshop(outputfile);

function runID(s){
    alert("InDesign running string returned from Photoshop: "  + s)
}

function runPhotoshop(f) {
    //open Photoshop and run psScript() function
    var bt = new BridgeTalk();  
    bt.target = "photoshop";  
    //a string representaion of variables and the function to run. 
    //a string var. Note string single, double quote syntax '"+s+"'
    bt.body = "var tempFolder = '"+f+"';" 
    
    //the function and the call
    bt.body += psScript(f).toString() + "psScript(f);";
    
    $.writeln(bt.body);

    bt.onResult = function(resObj) { 
        runID(resObj.body);
    }

    bt.onError = function( inBT ) { alert(inBT.body); };  
    bt.send(8);  
      
    function psScript(f) { 
        var myDoc = app.open(File(f));
        myDoc.ChangeMode(ChangeMode.GRAYSCALE);
		myDoc.save();
    }
}