Attribute VB_Name = "Module1"
Sub stockmkt()
Dim lastrow As Long

'Assuming date column in the dataset is in YYYYMMDD text format

'detect data range
lastrow = Range("A" & Rows.Count).End(xlUp).Row


'sort dataset
    Range("A1:E5").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Key2:=Range("B1") _
        , Order2:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:= _
        False, Orientation:=xlTopToBottom


 'initialize result set
    nd = 1 'output data row counter
    vol = 0 ' volume counter
    topen = 0 'ticket open variable
    tclose = 0 'ticket close variable
    pct = 0 ' percentage calc variable
        
'set up new column headers

    Range("I" & nd) = "Ticker"
    Range("J" & nd) = "Year"
    Range("K" & nd) = "Open"
    Range("L" & nd) = "Close"
    Range("M" & nd) = "Change"
    Range("N" & nd) = "Change%"
    Range("O" & nd) = "TotalVol"
    

'loop thru rows and create output rows
nd = nd + 1
topen = Range("C2").Value
For i = 2 To lastrow

            
            If Range("A" & i + 1) & Left(Range("B" & i + 1).Value, 4) <> Range("A" & i) & Left(Range("B" & i).Value, 4) Then
                
                               
                tclose = Range("F" & i).Value
                
                If topen = 0 Then
                pct = 0
                Else
                pct = (tclose - topen) / topen
                End If
                
                vol = vol + Range("G" & i).Value
                Range("I" & nd) = Range("A" & i)
                Range("J" & nd) = Left(Range("B" & i).Value, 4)
                Range("K" & nd) = topen
                Range("L" & nd) = tclose
                Range("M" & nd) = tclose - topen
                
               'format cell colour based on value
                If tclose - topen > 0 Then
                Range("M" & nd).Interior.ColorIndex = 10
                Else
                Range("M" & nd).Interior.ColorIndex = 3
                End If
                
                Range("N" & nd).NumberFormat = "0.00%" ' format cells to %
    
                Range("N" & nd) = pct
                Range("O" & nd) = vol
                nd = nd + 1
                vol = 0
                topen = Range("C" & i + 1).Value

            Else
                vol = vol + Range("G" & i).Value
                            
            End If
Next i



End Sub













