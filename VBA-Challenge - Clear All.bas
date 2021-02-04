Attribute VB_Name = "Module1"

Sub ClearAll()
    ' CLEAR ALL WORK IN EXCEL
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(LastRow, Columns.Count).End(xlToLeft).Column
        Range(ws.Cells(1, LastCol + 1), ws.Cells(LastRow + 1, LastCol + 99)).Clear
    Next ws
End Sub


Sub ClearBonus()
    For Each ws In Worksheets
    
        ws.Range("N1:ZZ7").Clear
    Next ws
End Sub
