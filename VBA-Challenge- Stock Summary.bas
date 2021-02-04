Attribute VB_Name = "Module1"
Sub Stock()
    For Each ws In Worksheets
        '' START OF SUMMARY TABLE
        ' STR is Summary Table Row to keep a counter and append to table
        Dim STR As Integer
        STR = 2
              
    
        ' Last Row in each Worksheet/last Col
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
    
        ' Array of headers
        Dim arrHeaders()
        arrHeaders = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    
        ' Inserting header after LastCol calculation
        For h = 0 To 3
            ws.Cells(1, LastCol + 2 + h) = arrHeaders(h)
        Next h
        
    
        ' Loop to read all tickers and append to Summary table
        Dim OpenPrice, StockVol, ClosePrice, PercentDelta As Double
        OpenPrice = ws.Range("C2").Value
        For i = 2 To LastRow
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
            ' If the Cell does not equal the next cell, we add all summary statistics
                ClosePrice = ws.Cells(i, 6).Value
                    ' If ClosePrice = 0 throws error so optionals
                    If (ClosePrice = 0) Then
                        If (OpenPrice > 0) Then
                            PercentDelta = -1
                        Else
                            PercentDelta = 0
                        End If
                    Else
                        'If Open Price = 0 to replace with new Open Price
                        If (OpenPrice = 0) Then
                            OpenPrice = ws.Cells(i + 1, 3).Value
                        End If
                        PercentDelta = Round((ClosePrice - OpenPrice) / OpenPrice, 4)
                    End If
                ws.Cells(STR, LastCol + 2).Value = ws.Cells(i, 1).Value
                ws.Cells(STR, LastCol + 3).Value = ClosePrice - OpenPrice
                
                    ' Conditional for highlighting Cell for Yearly Change
                    If (ClosePrice - OpenPrice < 0) Then
                        ws.Cells(STR, LastCol + 3).Interior.Color = vbRed
                    ElseIf (ClosePrice - OpenPrice > 0) Then
                        ws.Cells(STR, LastCol + 3).Interior.Color = vbGreen
                    End If
                    
                ws.Cells(STR, LastCol + 4).Value = PercentDelta
                ws.Cells(STR, LastCol + 5).Value = StockVol + ws.Cells(i, LastCol).Value
                StockVol = 0
                OpenPrice = ws.Cells(i + 1, 3).Value
                
                ' Next Row in summary stats
                STR = STR + 1
                
            Else
                ' Stock Volume addition
                StockVol = StockVol + ws.Cells(i, LastCol).Value
            End If
        
        Next i
        
        ' Conditional Formatting for Percent CHange column
        LastSumRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 2
        ws.Range(ws.Cells(2, LastCol + 4), ws.Cells(LastSumRow, LastCol + 4)).NumberFormat = "0.00%"
        
        '' END OF SUMMARY TABLE
        
    
        
    Next ws
  
 
End Sub
