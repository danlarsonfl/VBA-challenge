Sub StockTracker()


Dim StockTicker As String
Dim LastRow As Double
Dim LastTable As Double
Dim Yearly As Double
Dim StockVol As Double
Dim SummeryTable As Double
Dim OpenVal As Double
Dim CloseVal As Double

Dim WS As Worksheet
    For Each WS In ThisWorkbook.Worksheets

OpenVal = WS.Cells(2, 3).Value

SummeryTable = 2
StockVol = 0
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row


WS.Range("I1").Value = "Ticker"
WS.Range("J1").Value = "Yearly Change"
WS.Range("K1").Value = "Percent Change"
WS.Range("L1").Value = "Total Stock Volume"

WS.Range("N1").Value = "Greatest % Increase"
WS.Range("O1").Value = "Greatest % Decrease"
WS.Range("P1").Value = "Total Volume"

For i = 2 To LastRow

    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    StockTicker = WS.Cells(i, 1).Value
    WS.Range("I" & SummeryTable).Value = StockTicker
    

    StockVol = WS.Cells(i, 7).Value + StockVol
    WS.Range("L" & SummeryTable).Value = StockVol

    CloseVal = WS.Cells(i, 6).Value
    WS.Range("J" & SummeryTable).Value = CloseVal - OpenVal
    
    If OpenVal <> 0 Then
        WS.Range("K" & SummeryTable).Value = CloseVal / OpenVal
        Else
    End If

    SummeryTable = SummeryTable + 1
    StockVol = 0
    OpenVal = WS.Cells(i + 1, 3).Value

Else

    StockVol = WS.Cells(i, 7).Value + StockVol
    


End If

Next i

WS.Range("L:L").NumberFormat = "###,###"
WS.Range("J:J").Style = "Currency"
WS.Range("K:K").Style = "Percent"


LastTable = WS.Cells(Rows.Count, 10).End(xlUp).Row


For j = 2 To LastTable
    If WS.Cells(j, 10).Value >= 0 Then
        WS.Cells(j, 10).Interior.ColorIndex = 4
    Else
        WS.Cells(j, 10).Interior.ColorIndex = 3
    End If
Next j

'Dim Max As Double
'For k = 2 To LastTable
'    If WS.Cells(k, 10).Value > Max Then
'       Max = WS.Cells(k, 10).Value
'    Else
'    Max = WS.Cells(k, 10).Value
'        Range("N2") = Max
'    End If
'Next k



Next WS

End Sub



