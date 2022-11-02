# VBA-Challenge

I am adding my script here as I wasn't able to upload my excel workbook (file was too big) to the github. And Learning Assistant advised me to upload my script here. 

Thank you. 

Sub Real()

Dim Ticker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim TotalStock As Double
Dim openPrice As Double

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Range("J1").Value = "Ticker"
Range("K1").Value = "YearlyChange"
Range("L1").Value = "YearlyChangePercent"
Range("M1").Value = "TotalStockVolume"


'intialize variables
Counter = 2
TotalStock = Cells(2, 7)
openPrice = Cells(2, 3)

For i = 2 To Lastrow
    If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        Ticker = Cells(i, 1).Value
        YearlyChange = Cells(i, 6) - openPrice
        
        'display results
        Range("J" & Counter).Value = Ticker
        Range("K" & Counter).Value = YearlyChange
            If (YearlyChange < 0) Then
                Range("K" & Counter).Interior.ColorIndex = 3
            Else
                Range("K" & Counter).Interior.ColorIndex = 4
            End If
            
        Range("L" & Counter).Value = YearlyChange / openPrice
        Range("M" & Counter).Value = TotalStock
        
        'RE-assign
        TotalStock = Cells(i + 1, 7)
        openPrice = Cells(i + 1, 3)
        
        Counter = Counter + 1
    Else
        TotalStock = TotalStock + Cells(i + 1, 7)
    End If
Next i



End Sub
