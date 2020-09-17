Sub test()

'Loop through all worksheets
Dim WS As Worksheet
For Each WS In ActiveWorkbook.Worksheets
WS.Activate

'Declare new variables
Dim openPrice As Double
Dim closePrice As Double
Dim ticker As String
Dim yearlyChange As Double
Dim volume As Integer

volume = 0

'Add headers for main variables in summary table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Volume"

'Find the last row
lastRow = WS.Cells(Rows.Count, 1).End(xlUp).row

For i = 2 To lastRow
    ticker = Cells(i, 1).Value
    Cells(i, 9).Value = ticker
    
    'Set opening price
    openPrice = Cells(i, 3).Value
    
    'Set closing price
    closePrice = Cells(i, 6).Value
    
    'Calculate yearly change (openPrice - closePrice)
    yearlyChange = closePrice - openPrice
    Cells(i, 10) = yearlyChange
    
    'Calculate the percent change (yearlyChange / openPrice
     If (openPrice = 0 And closePrice = 0) Then
        percentChange = 0
    ElseIf (openPrice = 0 And closePrice <> 0) Then
        percentChange = 1
    Else
        percentChange = yearlyChange / openPrice
        Cells(i, 11).Value = percentChange
        Cells(i, 11).NumberFormat = "0.00%"
    End If
    
    'Calculate volume of stock
    volume = Cell(i, 7).Value
    Cells(i, 12).Value = volume
    
    
Next i

Next WS

End Sub
