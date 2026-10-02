' Copies Sheet1 once for every name listed in column A of Sheet2,
' and renames each copy to that name.
Sub CopySheet()

    Dim i As Long, LastRow As Long, ws As Worksheet

    Sheets("Sheet2").Activate

    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    For i = 1 To LastRow

        Sheets("Sheet1").Copy After:=Sheets(i)

        ActiveSheet.Name = Sheets("Sheet2").Cells(i, 1)

    Next i

End Sub
