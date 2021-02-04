Attribute VB_Name = "Module1"
Sub MaxVals()
    'Generate the Greatest %inc %dec and total vol for each worksheet
    Dim GVol, GInc, GDec As Double
    Dim GVolT, GincT, GDecT As String
    
    ' Goes through Worksheets
    For Each ws In Worksheets
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        LastRow = ws.Cells(Rows.Count, LastCol).End(xlUp).Row
        'Sets up toable for worksheet
        
        ws.Cells(1, LastCol + 4).Value = "Ticker"
        ws.Cells(1, LastCol + 5).Value = "Value"
        ws.Cells(2, LastCol + 3).Value = "Greatest % Increase"
        ws.Cells(3, LastCol + 3).Value = "Greatest % Decrease"
        ws.Cells(4, LastCol + 3).Value = "Greatest Total Volume"
        
        GVol = 0
        GInc = 0
        GDec = 0
        
        'Loops and checks each for greatest value. reassigns if value is greater than stored value
        For i = 2 To LastRow
            If (ws.Cells(i, LastCol) > GVol) Then
                GVol = ws.Cells(i, LastCol)
                GVolT = ws.Cells(i, LastCol - 3)
            End If
            If (ws.Cells(i, LastCol - 1) > GInc) Then
                GInc = ws.Cells(i, LastCol - 1)
                GincT = ws.Cells(i, LastCol - 3)
            End If
            If (ws.Cells(i, LastCol - 1) < GDec) Then
                GDec = ws.Cells(i, LastCol - 1)
                GDecT = ws.Cells(i, LastCol - 3)
            End If
        Next i
        
        'assigns values for greatest inc/dec/vol and ticker values
        ws.Cells(2, LastCol + 4) = GincT
        ws.Cells(2, LastCol + 5) = GInc
        ws.Cells(3, LastCol + 4) = GDecT
        ws.Cells(3, LastCol + 5) = GDec
        ws.Cells(4, LastCol + 4) = GVolT
        ws.Cells(4, LastCol + 5) = GVol
        Range(ws.Cells(2, LastCol + 5), ws.Cells(3, LastCol + 5)).NumberFormat = "0.00%"
        
    Next ws
        
    
End Sub
