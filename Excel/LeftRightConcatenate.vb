Sub LeftRightConcatenate()
    ' Data in A2 - Sample: GUNS_ROSES
    Dim Cell As Range
    Set aRange = Range("A2", Range("A" & Rows.Count).End(xlUp))
    ' Split at _underscore
    For Each Cell In aRange
        If InStr(Cell, "_") Then
            Cell.Offset(0, 1).Value = Split(Cell, "_")(0)
            Cell.Offset(0, 2).Value = Split(Cell, "_")(1)
        Else
            Cell.Offset(0, 1).Value = Cell
        End If
    Next
    Range("B1").Value2 = "LEFT"
    Range("C1").Value2 = "RIGHT"
    ' Concatenate with " and "
    Range("D1").Value2 = "CONCATENATE"
    For Each Cell In aRange
        If Not IsEmpty(Cell.Offset(0, 2)) Then
            Cell.Offset(0, 3).Value = _
            Cell.Offset(0, 1).Value & " and " & Cell.Offset(0, 2).Value
        Else
            Cell.Offset(0, 3).Value = _
            Cell.Offset(0, 1).Value & "" & Cell.Offset(0, 2).Value
        End If
    Next
    Cells.EntireColumn.AutoFit
End Sub
