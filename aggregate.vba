Sub tgr()

    Dim ws As Worksheet
    Dim wsDest As Worksheet

    Set wsDest = Sheets("All")

    For Each ws In ActiveWorkbook.Sheets
        If ws.Name <> wsDest.Name Then
            ws.Range("A1", ws.Range("A2").End(xlToRight).End(xlDown)).Copy
            wsDest.Cells(Rows.Count, "A").End(xlUp).Offset(2).PasteSpecial xlPasteValues
        End If
    Next ws

End Sub
