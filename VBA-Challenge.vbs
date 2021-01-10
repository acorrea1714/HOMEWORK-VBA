Sub Duo()
Call VBA_HOMEWORK
Call Bonus
Call Closing
End Sub
Function VBA_HOMEWORK()
'Set Variables
Dim ws As Worksheet
Dim OpenStock As Double
Dim CloseStock As Double
Dim Percent_Change As Double
Dim Yearly_Change As Double
Dim Run As Integer
Dim C As Long
Dim LastRow As Long
Dim Vol As Double
Dim Summary_Table_Row As Integer
    
Summary_Table_Row = 2
Run = 0
   
'For each worksheet copy the headers and excute nested loop
For Each ws In Worksheets
'Headers - (Header = define cell location = (&row,&column))
ws.Range("I1") = "Ticker"
ws.Range("J1") = "Yearly Change"
ws.Range("K1") = "Percent Change"
ws.Range("L1") = "Total Stock Volume"
'Count the number of rows in column
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Set iterator for loop
For C = 2 To LastRow
'If first cell in column 1 = second cell in column 1 then
If ws.Cells(C + 1, 1).Value = ws.Cells(C, 1).Value Then
'Set up counter - if true than = 1
Run = Run + 1
'Sum volume for all matching tickers = running total
Vol = Vol + ws.Cells(C, 7)
'If they match then set openstock value
If Run = 1 Then
OpenStock = ws.Cells(C, 3)
End If
'On last row
Else
'Add last volume to running total
Vol = Vol + ws.Cells(C, 7)
'In (&row(2),&column 9) print ticker symbol
ws.Cells(Summary_Table_Row, 9) = ws.Cells(C, 1)
'In (&row(2),&column 12) print new volume total
ws.Cells(Summary_Table_Row, 12) = Vol
'Set closestock value
CloseStock = ws.Cells(C, 6)
'If Open does not equal 0 then
If OpenStock <> 0 Then
'Calculate PercentChange
Percent_Change = ((CloseStock - OpenStock) / OpenStock)
'Calculate YearlyChange
Yearly_Change = CloseStock - OpenStock
'If Open = 0 then
Else
'Set Pct and Yr Change to 0 or you will get cant divided by 0 error.
Percent_Change = 0
Yearly_Change = 0
End If
'In (&row(2),&column 11) print PercentChange.
ws.Cells(Summary_Table_Row, 11) = Percent_Change
'Format PercentChange to percent
ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
ws.Cells(Summary_Table_Row, 12).NumberFormat = "000,000"
'In (&row(2),&column 10) print YearlyChange
ws.Cells(Summary_Table_Row, 10) = Yearly_Change
'If (&row(2),&column 10) has positive change then, format green.
If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
'If (&row(2),&column 10) does not have positive change, format red.
Else
ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
End If
'Reset running total Volume
Vol = 0
'Add 1 row to table row to print subsequent ticker,pctchg,yrchg,volume values
Summary_Table_Row = Summary_Table_Row + 1
'Reset loop counter
Run = 0
End If
Next C
'Reset print row
Summary_Table_Row = 2
'Move to next worksheet
Next ws
End Function
Function Bonus()
'Set Variables
Dim ws As Worksheet
Dim C As Long
Dim R As Integer
Dim Most As Double
Dim Least As Double
Dim LastRow As Long
Dim HiTick As String
Dim LoTick As String
Dim Summary_Table_Column As Integer
'Set Headers for each workset
For Each ws In Worksheets
ws.Range("o1") = "Ticker"
ws.Range("p1") = "Value"
ws.Range("n2") = "Greatest % Increase"
ws.Range("n3") = "Greatest % Decrease"
ws.Range("n4") = "Greatest Total Volume"
'Set Counters
Summary_Table_Column = 0
Most = 0
Least = 0
'Count Last Row
LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
'Set Iterator for Column Loop and Row Loop
For C = 11 To 12
For R = 2 To LastRow
If ws.Cells(R, C).Value > Most Then
Most = ws.Cells(R, C).Value
HiTick = ws.Cells(R, 9).Value
Else
End If
If ws.Cells(R, C).Value < Least Then
Least = ws.Cells(R, C).Value
LoTick = ws.Cells(R, 9).Value
Else
End If
Next R
ws.Cells(2 + Summary_Table_Column, 15).Value = HiTick
ws.Cells(2 + Summary_Table_Column, 16).Value = Most
ws.Cells(3 + Summary_Table_Column, 15).Value = LoTick
ws.Cells(3 + Summary_Table_Column, 16).Value = Least
ws.Cells(5, 16).Value = ""
HiTick = ""
LoTick = ""
Most = 0
Least = 0
Summary_Table_Column = Summary_Table_Column + 2
Next C
ws.Range("P2:P3").NumberFormat = "0.00%"
ws.Range("P4").NumberFormat = "000,00"
Next ws
End Function
Function Closing()
':)
MsgBox ("Have a great rest of your day! :)")
End Function
