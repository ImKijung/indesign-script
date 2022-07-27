var mydoc = app.activeDocument;
var docPath = mydoc.filePath;
var story_count = mydoc.stories.length;
var Report = new File(docPath +"/" + mydoc.name.replace(".indd","") + ".txt");
Report.open("w");
Report.encoding = "UTF-8";

for (var i=0 ; i<story_count ; i++)
{
    var paras = mydoc.stories[i].paragraphs;
    var para_count = paras.length;
    for (var j=0 ; j<para_count ; j++)
    {
        if (paras[j].appliedParagraphStyle.name.indexOf('Heading') != -1)
        {
            Report.writeln('<app order="' + j + '" file="' + mydoc.name.replace('.indd', '.idml') + '" depth="' + paras[j].appliedParagraphStyle.name + '" string="' + paras[j].contents.replace("\r","") + '"/>');
        }
    }
}
alert("완료, 인디자인 문서 폴더의 txt 파일을 확인하세요.");
Report.close();
Report.execute();