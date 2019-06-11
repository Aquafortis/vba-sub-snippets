Sub ColorIndexPrint()
    ' Print Index Colors to new first sheet for reference
    Sheets.Add(Before:=Sheets(1)).Name = "ColorIndex"
    Sheets("ColorIndex").Activate
    Dim cIndex As Integer
    For cIndex = 1 To 56
        Cells(cIndex, 1).Interior.ColorIndex = cIndex
        Cells(cIndex, 1) = cIndex
    Next
End Sub
