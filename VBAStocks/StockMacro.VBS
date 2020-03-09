Sub StockMacro()

'declare variables
Dim i, j, LastRow As Long
Dim Ticker As String
Dim Volume As Double
Dim TickerOpen As Double
Dim TickerClose As Double
Dim change As Double
Dim TickerTypes As Long
Dim ChangeMax As Double
Dim PercentMax As Double
Dim MaxVolume As Long
Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
With ws

'set last row
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row

'start TickerTypes counter at 0
TickerTypes = 0

'Add new headers
.Cells(1, 9).Value = ("Ticker")
.Cells(1, 10).Value = ("Yearly Change")
.Cells(1, 11).Value = ("Percent Change")
.Cells(1, 12).Value = ("Total Stock Volume")

'Begin primary For loop
For i = 2 To LastRow
'error handling to deal with missing stock data
On Error Resume Next
'When row is the same as the row below, adds up the volume
    If .Cells(i + 1, 1).Value = .Cells(i, 1) Then
        Volume = Volume + .Cells(i, 7)
        'Check for first row of the Ticker by checking Ticker value above, then sets the opening value
        If .Cells(i - 1, 1).Value <> .Cells(i, 1) Then
            TickerOpen = .Cells(i, 3)
        End If
'When row is different from the row below it (aka the final row for a ticker)
    ElseIf .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
        'final volume added
        Volume = Volume + .Cells(i, 7)
        'add 1 to the ticker types counter
        TickerTypes = 1 + TickerTypes
        'the close value is set
        TickerClose = .Cells(i, 6)
        'calculates the yearly change and reports out
        .Cells(TickerTypes + 1, 10) = TickerClose - TickerOpen
        'sets percent cell to % formattings and reports out the percent change
        .Cells(TickerTypes + 1, 11).NumberFormat = "0.00%"
        .Cells(TickerTypes + 1, 11) = .Cells(TickerTypes + 1, 10) / TickerOpen
        'reports out the ticker
        .Cells(TickerTypes + 1, 9) = .Cells(i, 1).Value
        'reports out the final volume
        .Cells(TickerTypes + 1, 12) = Volume
        'conditional formatting of the yearly change
            If .Cells(TickerTypes + 1, 10) > 0 Then
                .Cells(TickerTypes + 1, 10).Interior.ColorIndex = 4
            ElseIf .Cells(TickerTypes + 1, 10) < 0 Then
                .Cells(TickerTypes + 1, 10).Interior.ColorIndex = 3
            End If
        'resets the ticker open, close, and volume at zero
        TickerClose = 0
        TickerOpen = 0
        Volume = 0
    End If
Next i

'find the greatest percent increase and return the value, ticker associated, and a row title
For i = 2 To (TickerTypes + 1)
    If .Cells(i, 11) = Application.WorksheetFunction.Max(.Range("K2:K" & (TickerTypes + 1))) Then
      .Cells(2, 15).Value = "Greatest % Increase"
      .Cells(2, 17).NumberFormat = "0.00%"
      .Cells(2, 16).Value = .Cells(i, 9).Value
      .Cells(2, 17).Value = .Cells(i, 11).Value

    End If
Next i
'find greatest percent decrease and return the value, the ticker associated, and a row title
For i = 2 To (TickerTypes + 1)
    If .Cells(i, 11) = Application.WorksheetFunction.Min(.Range("K2:K" & (TickerTypes + 1))) Then
      .Cells(3, 15).Value = "Greatest % Decrease"
      .Cells(3, 17).NumberFormat = "0.00%"
      .Cells(3, 16).Value = .Cells(i, 9).Value
      .Cells(3, 17).Value = .Cells(i, 11).Value

    End If
Next i
'find the greatest total volume and return the value, title associated, and a row title
For i = 2 To (TickerTypes + 1)
    If .Cells(i, 12) = Application.WorksheetFunction.Max(.Range("L2:L" & (TickerTypes + 1))) Then
      .Cells(4, 15).Value = "Greatest Total Volume"
      .Cells(4, 16).Value = .Cells(i, 9).Value
      .Cells(4, 17).Value = .Cells(i, 12).Value

    End If
Next i

'add headers to the 3 max value categories
.Cells(1, 16).Value = ("Ticker")
.Cells(1, 17).Value = ("Value")

End With
Next

End Sub
