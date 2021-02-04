Attribute VB_Name = "Module1"
Sub GOAT()
    '' GENERATES THE GREATEST STOCK VOL INCREASE AND DECREASE OF ALL! Worksheets
    ' Initializing variables including holding vars
    Dim MaxVolWs(), GreatPercIncWs(), GreatPercDecWs() As Variant
    Dim MaxVolTick(), GreatPercIncTick(), GreatPercDecTick() As String
    Dim CountVol, CountInc, CountDec As Integer
    ' starting loop to read all relavant columns
    ReDim MaxVolWs(1), GreatPercIncWs(1), GreatPercDecWs(1)
    ReDim MaxVolTick(1), GreatPercIncTick(1), GreatPercDecTick(1)
    MaxVolWs(0) = 0
    GreatPercIncWs(0) = 0
    GreatPercDecWs(0) = 0
    
    For Each ws In Worksheets
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        For i = 2 To LastRow
        
            ' Maximum Vol for all the Worksheets into an array (incase multiple)
            If (ws.Cells(i, LastCol) > MaxVolWs(0)) Then
                Erase MaxVolWs
                Erase MaxVolTick
                CountVol = 0
                ReDim MaxVolWs(0), MaxVolTick(0)
                MaxVolWs(0) = ws.Cells(i, LastCol)
                MaxVolTick(0) = ws.Cells(i, LastCol - 3)
            ElseIf (ws.Cells(i, LastCol) = MaxVolWs(0)) Then
                CountVol = CountVol + 1
                ReDim MaxVolWs(CountVol), MaxVolTick(CountVol)
                MaxVolWs(CountVol) = ws.Cells(i, LastCol)
                MaxVolTick(CountVol) = ws.Cells(i, LastCol - 3)
            End If
            
            'Greatest Increase
            If (ws.Cells(i, LastCol - 1) > GreatPercIncWs(0)) Then
                Erase GreatPercIncWs
                Erase GreatPercIncTick
                CountInc = 0
                ReDim GreatPercIncWs(0), GreatPercIncTick(0)
                GreatPercIncWs(CountInc) = ws.Cells(i, LastCol - 1)
                GreatPercIncTick(CountInc) = ws.Cells(i, LastCol - 3)
            ElseIf (ws.Cells(i, LastCol - 1) = GreatPercIncWs(0)) Then
                CountInc = CountInc + 1
                ReDim GreatPercIncWs(CountInc), GreatPercIncTick(CountInc)
                GreatPercIncWs(CountInc) = ws.Cells(i, LastCol - 1)
                GreatPercIncTick(CountInc) = ws.Cells(i, LastCol - 3)
            End If
            
            'Greatest Decrease
            If (ws.Cells(i, LastCol - 1) < GreatPercDecWs(0)) Then
                Erase GreatPercDecWs
                Erase GreatPercDecTick
                CountDec = 0
                ReDim GreatPercDecWs(0), GreatPercDecTick(0)
                GreatPercDecWs(CountDec) = ws.Cells(i, LastCol - 1)
                GreatPercDecTick(CountDec) = ws.Cells(i, LastCol - 3)
            ElseIf (ws.Cells(i, LastCol - 1) = GreatPercDecWs(0)) Then
                CountDec = CountDec + 1
                ReDim GreatPercDecWs(CountDec), GreatPercDecTick(CountDec)
                GreatPercDecWs(CountDec) = ws.Cells(i, LastCol - 1)
                GreatPercDecTick(CountDec) = ws.Cells(i, LastCol - 3)
            End If

        Next i
    Next ws
                
            
        
        ' Formatting for greatest% inc/dec and vol of all tables
    ActiveWorkbook.Worksheets(1).Activate
    LastCol2 = Cells(1, Columns.Count).End(xlToLeft).Column
    Cells(1, LastCol2 + 4).Value = "Ticker"
    Cells(1, LastCol2 + 5).Value = "Value"
    Cells(2, LastCol2 + 3).Value = "Greatest % Increase"
    Cells(2, LastCol2 + 4).Value = Join(GreatPercIncTick, ", ")
    Cells(2, LastCol2 + 5).Value = Join(GreatPercIncWs, ", ")
    Cells(3, LastCol2 + 3).Value = "Greatest % Decrease"
    Cells(3, LastCol2 + 4).Value = Join(GreatPercDecTick, ", ")
    Cells(3, LastCol2 + 5).Value = Join(GreatPercDecWs, ", ")
    Cells(4, LastCol2 + 3).Value = "Greatest Total Volume"
    Cells(4, LastCol2 + 4).Value = Join(MaxVolTick, ", ")
    Cells(4, LastCol2 + 5).Value = Join(MaxVolWs, ", ")
    
    Range(Cells(2, LastCol2 + 5), Cells(3, LastCol2 + 5)).NumberFormat = "0.00%"
    
    
End Sub