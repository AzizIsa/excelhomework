
'*******************************Task 2 (Moderate)*******************************
Sub stockChange():

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

'to get a name value of the worksheet
WorksheetName = ws.Name

SummaryTableRow = 2

    ws.Columns("J:J").Insert Shift:=xlToRight
    ws.Columns("J:J").Insert Shift:=xlToRight
    
ws.Range("J1").Value = "YearlyChange"
ws.Range("K1").Value = "PercentChange"

lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

On Error GoTo ErrorHandler: 'TO avoid errors from division by zero

For i = 2 To lastrow


stockTicker = ws.Cells(i, 9).Value
opendate = WorksheetName & "0101"
closedate = WorksheetName & "1230"
openPrice = WorksheetFunction.VLookup(stockTicker, ws.Range("A:C"), 3, False)
closePrice = WorksheetFunction.VLookup(stockTicker & "|" & closedate, ws.Range("A:F"), 6, True)

    yChange = closePrice - openPrice
    
  
    pChange = (yChange / openPrice)

    ws.Range("J" & SummaryTableRow).Value = yChange
    ws.Range("K" & SummaryTableRow).Value = pChange
    
    SummaryTableRow = SummaryTableRow + 1
    
ErrorHandler:
    If Err.Number = 11 Then
        ws.Cells(i, 11).Value = 0
        Resume AfterError:
    End If
AfterError:

Next i

Next ws

End Sub

