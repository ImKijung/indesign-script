Dim usedRange, strXML
Dim app, book, sheet

Main()
returnValue = strXML


Sub Main() 
    Set app = CreateObject("Excel.Application")
    app.visible = false
    Set book = app.WorkBooks.Open(arguments(0))
    Set sheet = book.WorkSheets(1)
    ExportRangeToXML()
    Set sheet = nothing
    Set book = nothing
    app.ActiveWorkbook.Close
    app.Quit 
    Set app = nothing
End Sub


Sub ExportRangeToXML()
 
Dim varTable
Dim intRow
Dim intCol
Dim intFileNum
Dim strFilePath
Dim strRowElementName
Dim strTableElementName
Dim varColumnHeaders
 
    strTableElementName = "Table"
    strRowElementName = "Row"
 
    Set usedRange = sheet.UsedRange
    varTable =usedRange.Value
    varColumnHeaders = usedRange.Rows(1).Value
 
    strXML = "<?xml version=""1.0"" encoding=""utf-8""?>"
    strXML = strXML & "<" & strTableElementName & ">"
    For intRow = 2 To UBound(varTable, 1)
        strXML = strXML & "<" & strRowElementName & " imgName=""" & varTable(intRow, 3) & """>"
        For intCol = 4 To UBound(varTable, 2)
            strXML = strXML & "<" & varColumnHeaders(1, intCol) & ">" & _
                varTable(intRow, intCol) & "</" & varColumnHeaders(1, intCol) & ">"
        Next
        strXML = strXML & "</" & strRowElementName & ">"
    Next
    strXML = strXML & "</" & strTableElementName & ">"
 
End Sub

