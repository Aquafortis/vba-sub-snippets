Sub MergeAllSheets()
    ' Merge all sheets into one
    ' Keep top row of first sheet only
    Dim S As Integer
    Dim cRng As Range
    Sheets.Add(Before:=Sheets(1)).Name = "Merged"
    Sheets(2).Range("A1").EntireRow.Copy _
    Sheets(1).Range("A1")
    For S = 2 To Sheets.Count
        Sheets(S).Activate
        Set cRng = ActiveSheet.UsedRange
        cRng.Offset(1, 0).Resize(cRng.Rows.Count - 1, _
        cRng.Columns.Count).Copy _
        Sheets(1).Range("A65536").End(xlUp)(2)
    Next
    Sheets("Merged").Activate
End Sub
