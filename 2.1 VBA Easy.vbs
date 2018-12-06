'********************Task 1 (EASY)********************
Sub StockData():

For Each ws In Worksheets

'Declare Variables

Dim stockName As String
Dim volTotal As Double
Dim SummaryTableRow As Integer

'Assign Values to Variables

volTotal = 0
SummaryTableRow = 2

'Count the last row of the first column of the worksheet

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Add ColumnNames for output data

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Total_Volume"


'Loop through all stock tickers

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

    ticker = ws.Cells(i, 1).Value
    volTotal = volTotal + ws.Cells(i, 7).Value

    ws.Range("I" & SummaryTableRow).Value = ticker
    ws.Range("J" & SummaryTableRow).Value = volTotal

    SummaryTableRow = SummaryTableRow + 1
    volTotal = 0
Else

    volTotal = volTotal + ws.Cells(i, 7).Value
    
End If

Next i

Next ws

End Sub



