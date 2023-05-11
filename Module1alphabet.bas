Attribute VB_Name = "Module1"

Sub vba_challenge_alphabetic()

'loop through all work sheets
For Each ws In Worksheets

'now we create variables for ticker,yearly change,yearly percentage change, and total volume for each ticker
Dim ticker As String
Dim yrchange, percentage As Double
Dim tvolume, Lastrow As Long

'to keep the track of the location for each ticker in other table
Dim table2row As Integer

table2row = 2

'determine the last row
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'lets name columns with ticker,yearly change, percentage change,and total stock volume
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percentage Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

'we add variables for yearly open price and yearly close price
Dim yro, yrc As Double

'assigning X as counter
Dim x As Long
x = 0

'loop through all tickers
For i = 2 To Lastrow

'check if we are still in the same ticker,if we aren't then finalize the process and start with another ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'we give data to yearly open
yro = ws.Cells(i - x, 3).Value

'set ticker name
ticker = ws.Cells(i, 1).Value

'add total volume
tvolume = tvolume + ws.Cells(i, 7).Value

'now finalize yearly change
yrc = ws.Cells(i, 6).Value

yrchange = yrc - yro

'calculating the percentage of yearly change to the openning price
percentage = yrchange / yro

'now we print ticker's name into other column
ws.Range("i" & table2row).Value = ticker

'now we print yearly change
ws.Range("j" & table2row).Value = yrchange

'now we print percentage of yearly change
ws.Range("k" & table2row).Value = percentage
ws.Range("k" & table2row).NumberFormat = "0.00%"

'now lets print total volume
ws.Range("L" & table2row).Value = tvolume

'add one row to the  new columns to start with another ticker
table2row = table2row + 1

'reset total volume before we start with another ticker
tvolume = 0

'reset rows counter
x = 0

'2nd condition which is if we are still in the same ticker data
Else

'X the counter of rows should go up
x = x + 1

'we add data to total volume
tvolume = tvolume + ws.Cells(i, 7).Value

End If
Next i

'now we format yearly change according to their values
Dim lastrow2 As Long

lastrow2 = ws.Cells(Rows.Count, 10).End(xlUp).Row

'loop through all cells in that column
For j = 2 To lastrow2

'conditions to color cells according to their value
'color it red if its value is below zero
If ws.Cells(j, 10).Value < 0 Then
    
    ws.Cells(j, 10).Interior.ColorIndex = 3
    
'otherwise green
Else
    ws.Cells(j, 10).Interior.ColorIndex = 4

End If
Next j


'Now find the greatest percentage increase, greatest percentage decrease (the lowest percentage) and the greatest volume
'first off we create variables for greatest, lowest increase in percentage and the greatest total volume
Dim Hincrease, Lincrease As Double
Dim HTVolume, Lastrow3 As Long

'lastrow3 to count rows numbers in column "i" will be used later
Lastrow3 = ws.Cells(Rows.Count, 9).End(xlUp).Row

'now we name new rows
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'we name new columns
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

'lets find out the greatest percentage increase
Hincrease = Application.Max(ws.Range("K2:K" & Lastrow3).Value)

'next finding out the greatest percentage decrease (the lowest number)
Lincrease = Application.Min(ws.Range("K2:K" & Lastrow3).Value)

'now finding out the greatest total volume
HTVolume = Application.Max(ws.Range("L2:L" & Lastrow3).Value)

'now we loop through tickers in column i

For v = 2 To Lastrow3

'now search and match values to retrieve other data and print then in other rows/columns
'1st condition to search and match the greatest increase
If ws.Cells(v, 11).Value = Hincrease Then
    
    ws.Cells(2, 16).Value = ws.Cells(v, 9).Value
    ws.Cells(2, 17).Value = ws.Cells(v, 11).Value
    ws.Cells(2, 17).NumberFormat = "0.00%"

'2nd condition for the greatet decrease
ElseIf ws.Cells(v, 11).Value = Lincrease Then

    ws.Cells(3, 16).Value = ws.Cells(v, 9).Value
    ws.Cells(3, 17).Value = ws.Cells(v, 11).Value
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
'3rd condition for the greatest total volume
ElseIf ws.Cells(v, 12).Value = HTVolume Then
    
    ws.Cells(4, 16).Value = ws.Cells(v, 9).Value
    ws.Cells(4, 17).Value = ws.Cells(v, 12).Value

End If
Next v

Next ws

End Sub
