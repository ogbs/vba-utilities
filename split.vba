Sub CreateNewWBS()
Dim wbThis As Workbook
Dim wbNew As Workbook
Dim ws As Worksheet
Dim strFilename As String

Set wbThis = ThisWorkbook
For Each ws In wbThis.Worksheets
    strFilename = wbThis.Path & "/" & ws.Name
    ws.Copy
    Set wbNew = ActiveWorkbook
    wbNew.SaveAs strFilename, FileFormat:=xlCSV, CreateBackup:=False
    wbNew.Close True
Next ws
End Sub
