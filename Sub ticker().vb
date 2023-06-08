Sub ticker()

'initial variable for holding ticker
Dim ticker_name As String

'var for holding opening price
Dim oprice As Double

'var for holding closing price
Dim cprice As Double

'var for holding yearly change
Dim year_change As Double

'var for percent change from open/close price
Dim per_change As Double

'var for total stock volume
Dim total_vol As Double

'keep track of location for each ticker in summary table
Dim strow As Integer

'var for looping through
Dim i, j As Integer
Dim LastRow As Double

'var for ws
Dim ws As Worksheet

'var for greatest % increase/decrease and greatest total volume
Dim greatest, lowest, gtv As Double

'loop through all data
For Each ws In Worksheets

'print all new headers / add this bw ForEach ws in Worksheets later
ws.Cells(1, 9).Value = "ticker"
ws.Cells(1, 16).Value = "ticker"
ws.Cells(1, 10).Value = "yearly change"
ws.Cells(1, 11).Value = "percent change"
ws.Cells(1, 12).Value = "total stock volume"
ws.Cells(1, 17).Value = "value"

'last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'change width of columns
ws.Columns("I:L").ColumnWidth = 17
ws.Columns("O").ColumnWidth = 20
ws.Columns("Q").ColumnWidth = 10

strow = 2
'set opening price
    oprice = ws.Cells(2, 3).Value
'set greatest and lowest to something that will get overwritten
greatest = -10 * 1000000
lowest = 10 * 1000000
'same with gtv
gtv = -10 * 1000000
        
For i = 2 To LastRow

'check to see same ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'set ticker name
    ticker_name = ws.Cells(i, 1).Value
        
'add total
    total_vol = total_vol + ws.Cells(i, 7).Value
        
'print ticker name
    ws.Range("I" & strow).Value = ticker_name
        
'print total
    ws.Range("L" & strow).Value = total_vol
        
'new close
    cprice = ws.Cells(i, 6).Value
        
'print yearly change
    year_change = cprice - oprice
    ws.Range("J" & strow).Value = year_change
        
'change color of cell based on +/-
        If year_change > 0 Then
            ws.Range("J" & strow).Interior.ColorIndex = 4
        ElseIf year_change < 0 Then
            ws.Range("J" & strow).Interior.ColorIndex = 3
        Else
            ws.Range("J" & strow).Interior.ColorIndex = 0
        End If
'print percent change
    per_change = year_change / oprice
    
'per_change = Format("0.00%")

    ws.Range("K" & strow).Value = per_change
        
        If ws.Range("K" & strow).Value > greatest Then
            greatest = ws.Cells(strow, 11).Value
        ElseIf ws.Range("K" & strow).Value < lowest Then
            lowest = ws.Cells(strow, 11).Value
        End If
    
'gtv
        If total_vol > gtv Then
            gtv = total_vol
            
        End If
        
'new open
    oprice = ws.Cells(i + 1, 3).Value
    
'add one to sumtablerow
    strow = strow + 1
        
'reset total
    total_vol = 0
    
'if same brand, just add to total
ElseIf ws.Cells(i + 1, 1).Value = ws.Cells(i, 1).Value Then
        total_vol = total_vol + ws.Cells(i, 7).Value
        
End If
    
'finish out
Next i

'correct format
'ws.Range("K2:K" & LastRow).NumberFormat = "0.00%"

'new table
ws.Cells(2, 15).Value = "Greatest % increase"
ws.Cells(3, 15).Value = "Greatest % decrease"
ws.Cells(4, 15).Value = "Greatest Total volume"

'print greatest,lowest
ws.Cells(2, 17).Value = greatest
ws.Cells(3, 17).Value = lowest
ws.Cells(4, 17).Value = gtv

'did not know how else to do this part, tried v/hlookup and tooltip lead me to xlookup
ws.Cells(2, 16).Value = WorksheetFunction.XLookup(ws.Cells(2, 17).Value, ws.Range("K2:K" & LastRow).Value, ws.Range("I2:I" & LastRow).Value)
ws.Cells(3, 16).Value = WorksheetFunction.XLookup(ws.Cells(3, 17).Value, ws.Range("K2:K" & LastRow).Value, ws.Range("I2:I" & LastRow).Value)
ws.Cells(4, 16).Value = WorksheetFunction.XLookup(ws.Cells(4, 17).Value, ws.Range("L2:L" & LastRow).Value, ws.Range("I2:I" & LastRow).Value)

'add % sign
ws.Cells(2, 17).Value = greatest * 100 & "%"
ws.Cells(3, 17).Value = lowest * 100 & "%"

Next ws

End Sub
