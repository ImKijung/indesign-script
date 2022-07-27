#targetengine "session";

if (app.documents == 0) {
    alert("인디자인 문서를 열어놓은 상태에서 실행하세요.");
    exit();
}

var win = new Window ("palette", "New 단락스타일 컨텐츠 추출하기");
var maingroup = win.add ("panel {orientation: 'column'}");
    add_row (maingroup);
var show_btn = win.add ("button", undefined, "show");

show_btn.onClick = function () {
    var txt = " ";
    var myErr = mySucc = 0;    
    var myList = "";
    var doc = app.activeDocument;
    
    var docPath = doc.filePath;
    var Report = new File(docPath +"/" + doc.name.replace(/\.indd$/,"") + "_ps_contents.txt");
    Report.encoding = "UTF-8"; //Encoding for Windows/Excel
    Report.open("w");
    
    for (var n = 0; n < maingroup.children.length; n++) {
        txt += maingroup.children[n].list.selection.text + "\n";
        ps_name = maingroup.children[n].list.selection.text;
        // $.writeln(ps_name);
        print_contents(ps_name, Report);
    }
    
    Report.close();
    Report.execute();
    //alert ("Rows: \n" + txt);
}
win.show ();

function add_row (maingroup) {
    var group = maingroup.add ("group");
    group.list = group.add ("dropdownlist", [" ", " ", 200, 20], listItems);
    var doc = app.activeDocument;
    var paraStyles = doc.allParagraphStyles;
    var listItems = [];
    for( var l = 0; l < paraStyles.length; l++ ) {
        var listItems = group.list.add('item', paraStyles[l].name);
    }
    group.plus = group.add ("button", undefined, "+");
    group.plus.onClick = add_btn;
    group.minus = group.add ("button", undefined, "-");
    group.minus.onClick = minus_btn;
    group.index = maingroup.children.length - 1;
    win.layout.layout (true);
}
function add_btn () {
    add_row (maingroup);
}
function minus_btn () {
    maingroup.remove (this.parent);
    win.layout.layout (true);
}

function print_contents (ps_name, Report) {
    var doc = app.activeDocument;
    // var ps_contents = new Array()
    
    app.findTextPreferences = NothingEnum.nothing; 
    app.changeTextPreferences = NothingEnum.nothing;
    app.findTextPreferences.appliedParagraphStyle = ps_name;

    var myFound = doc.findText();
    for (i = 0; i < myFound.length; i ++) {
        var myContent = myFound[i].contents;
        var myPage = myFound[i].parentTextFrames[0].parentPage.name;
        
        var ps_contents = myPage + "\t" + ps_name + "\t" + myContent.split('\r') + "\r" ;
        // $.writeln(ps_contents);
        Report.writeln(ps_contents.split(",").join(""));
        app.changeTextPreferences = NothingEnum.nothing;
        app.findTextPreferences = NothingEnum.nothing;
    }
}