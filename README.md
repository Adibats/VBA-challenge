# VBA-challenge.

3 separate methods, must be run in order. First method populates all tickers. Second method populates values for each individual ticker. Third method retrieves top results from list of tickers.


Public Sub getTickers()
Dim current As Worksheet

For Each ws In Worksheets

Dim endRow As Long
Dim ticker As String
Dim j As Long

j = 3
ticker = ws.Range("A2").Value
ws.Range("i1").Value = "Ticker"
ws.Range("i2").Value = ticker

endRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To endRow

    If ws.Cells(i, 1) <> ticker Then
        ws.Cells(j, 9).Value = ws.Cells(i, 1).Value
        ticker = ws.Cells(i, 1).Value
        j = j + 1
    End If

Next i
Next ws
End Sub

Public Sub getMetrics()

Dim current As Worksheet

For Each ws In Worksheets


Dim endRow As Long
Dim ticker As String
Dim endDate As Double
Dim begDate As Double
Dim j As Long
Dim volume As LongLong

ws.Range("j1").Value = "Yearly Change"
ws.Range("k1").Value = "Percent Change"
ws.Range("l1").Value = "Total Stock Volume"

ticker = ws.Range("i2").Value
begDate = 2
endDate = begDate
j = 2
volume = 0
endRow = Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To endRow
    If ws.Cells(i, 1).Value = ticker Then
        If CLng(ws.Cells(i + 1, 2).Value) >= CLng(ws.Cells(i, 2).Value) Then
            endDate = i + 1
            ws.Cells(j, 10).Value = ws.Cells(endDate, 6).Value - ws.Cells(begDate, 3).Value
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            ws.Cells(j, 11).Value = FormatPercent(ws.Cells(endDate, 6).Value / ws.Cells(begDate, 3).Value - 1)
            volume = volume + ws.Cells(i, 7).Value
            ws.Cells(j, 12).Value = volume + ws.Cells(i + 1, 7)
        End If
    Else
        ticker = ws.Cells(i, 1).Value
        j = j + 1
        begDate = i
        endDate = begDate
        volume = ws.Cells(i, 7).Value
    End If
        
Next i

Next ws

End Sub

Public Sub Totals()
Dim ws As Worksheet

For Each ws In Worksheets
Dim endRow As LongLong
Dim increase As Double
Dim decrease As Double
Dim volume As LongLong
Dim i As LongLong

ws.Range("p1").Value = "Ticker"
ws.Range("q1").Value = "Value"
ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greatest Total Volume"


endRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
increase = ws.Cells(2, 11).Value
decrease = increase
volume = ws.Cells(2, 12).Value
ws.Range("q2").Value = FormatPercent(increase)
ws.Range("p2").Value = ws.Cells(2, 9).Value
ws.Range("q3").Value = FormatPercent(decrease)
ws.Range("p3").Value = ws.Cells(2, 9).Value
ws.Range("p4").Value = ws.Cells(2, 9).Value
        
For i = 2 To endRow
    If ws.Cells(i, 11).Value > increase Then
        increase = ws.Cells(i, 11).Value
        ws.Range("q2").Value = FormatPercent(increase)
        ws.Range("p2").Value = ws.Cells(i, 9).Value
    
    ElseIf ws.Cells(i, 11).Value < decrease Then
        decrease = ws.Cells(i, 11).Value
        ws.Range("q3").Value = FormatPercent(decrease)
        ws.Range("p3").Value = ws.Cells(i, 9).Value
    End If
    If ws.Cells(i, 12).Value > volume Then
        volume = ws.Cells(i, 12).Value
        ws.Range("q4").Value = volume
        ws.Range("p4").Value = ws.Cells(i, 9).Value
    End If

Next i
Next ws
End Sub


