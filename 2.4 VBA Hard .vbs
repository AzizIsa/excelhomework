'****************************Task 3 (Hard)*******************************
Sub Greatest_Results():
'Perform the formatting changes for all existing worksheets.
For Each ws In Worksheets

' Declaring variables with data types

Dim maxTicker As String
Dim minTicker As String
Dim totalTicker As String
Dim maxValues As Double
Dim minValues As Double
Dim totalVol As Double

' creating parameterNames(columns names) for the values

ws.Range("O1").Value = "Ticker"
ws.Range("P1").Value = "Value"
ws.Range("N2").Value = "Greatest % Increase"
ws.Range("N3").Value = "Greatest % Decrease"
ws.Range("N4").Value = "Greatest Total Volume"

' Calculating Max, Min, total values from the given ranges.

maxValues = WorksheetFunction.Max(ws.Range("K:K"))
minValues = WorksheetFunction.Min(ws.Range("K:K"))
totalVol = WorksheetFunction.Max(ws.Range("L:L"))

 ws.Range("P2").Value = maxValues
 ws.Range("P3").Value = minValues
 ws.Range("P4").Value = totalVol
 
 lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
 
 
'Looping through to vLookup tickers based on the min,max, total values above

 For i = 2 To lastrow
 
 If ws.Cells(i, 11).Value = maxValues Then
 
 maxTicker = ws.Cells(i, 9).Value
 ws.Range("O2").Value = maxTicker
 
ElseIf ws.Cells(i, 11).Value = minValues Then
 
    minTicker = ws.Cells(i, 9).Value
    ws.Range("O3").Value = minTicker
    
ElseIf ws.Cells(i, 12).Value = totalVol Then

totalTicker = ws.Cells(i, 9).Value
ws.Range("O4").Value = totalTicker

End If
 
 Next i
 
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"
ws.Columns("N:P").AutoFit
Next ws

End Sub