var docs = app.documents;

var pdfPreset = app.pdfExportPresets.item("[고품질 인쇄]");

app.pdfExportPreferences.viewPDF = false; // PDF 출력 후 열기 옵션 끔

for (var i=0;i<docs.length;i++) {
  var docPath = docs[i].filePath;
  var docName = docs[i].name;
  var exportName = docName.replace('.indd', '.pdf');
  var output = new File(docPath + '/' + exportName);
  docs[i].exportFile(ExportFormat.pdfType, output, false, pdfPreset);
}

app.pdfExportPreferences.viewPDF = true; // PDF 출력 후 열기 옵션 켬
alert(exportName);
