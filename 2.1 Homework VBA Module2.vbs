
'**********Task 2 (Hard)**********
Sub stockChange():

On Error Resume Next 'TO avoid errors from division by zero

For Each ws In Worksheets
'Insert 2 Columns to the left of Column J
Dim WorksheetName As String
Dim opendate As String
Dim closedate As String
Dim stockTicker As String
Dim openPrice As Double
Dim closePrice As Double
Dim yChange As Double
Dim SummaryTableRow As Integer
Dim pChange As Double


WorksheetName = ws.Name

SummaryTableRow = 2

    ws.Columns("J:J").Insert Shift:=xlToRight
    ws.Columns("J:J").Insert Shift:=xlToRight
    
ws.Range("J1").Value = "YearlyChange"
ws.Range("K1").Value = "PercentChange"

lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow


stockTicker = ws.Cells(i, 9).Value
opendate = WorksheetName & "0101"
closedate = WorksheetName & "1230"
openPrice = WorksheetFunction.VLookup(stockTicker, ws.Range("A:C"), 3, False)
closePrice = WorksheetFunction.VLookup(stockTicker & "|" & closedate, ws.Range("A:F"), 6, True)

    yChange = closePrice - openPrice
    
  
    pChange = (closePrice - openPrice) / openPrice

    ws.Range("J" & SummaryTableRow).Value = yChange
    ws.Range("K" & SummaryTableRow).Value = pChange
    
    SummaryTableRow = SummaryTableRow + 1

Next i

Next ws

End Sub

Sub Conditional_FOrmating():
For Each ws In Worksheets



lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastrow

If ws.Cells(i, 10).Value >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

Else

ws.Cells(i, 10).Interior.ColorIndex = 3

End If


Next i

ws.Range("J1").EntireColumn.NumberFormat = "0.000000000"
ws.Range("K1").EntireColumn.NumberFormat = "0.00%"
ws.Range("P2").NumberFormat = "0.00%"
ws.Range("P3").NumberFormat = "0.00%"
ws.Columns("I:L").AutoFit

Next ws
End Sub

Sub Greatest_Results():

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

ws.Columns("N:P").AutoFit
Next ws

End Sub





