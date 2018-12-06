'****************************Task 2 (Moderate, Formatting)*******************************
Sub Conditional_FOrmating():

'Perform the formatting changes for all existing worksheets.

For Each ws In Worksheets
'Count last row of the column 9

lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'Loop through the vales in the column 9 and apply conditional formatting (Red/green )

For i = 2 To lastrow

If ws.Cells(i, 10).Value >= 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

Else
ws.Cells(i, 10).Interior.ColorIndex = 3

End If

Next i


'Apply appropriate formating to match the results in the screenshots.
ws.Range("J1").EntireColumn.NumberFormat = "0.000000000"
ws.Range("K1").EntireColumn.NumberFormat = "0.00%"

'Autofit columns so that info is fully readable.
ws.Columns("I:L").AutoFit

Next ws
End Sub