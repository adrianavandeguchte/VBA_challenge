Sub StockMacro()

'this macro is intended for worksheets with yearly stock data
'data must be sorted by date per unique ticker
'for each unique ticker the yearly change, percent change, and yearly volume moved will be outputted to a summery table
'from each summery table the ticker coresponding to the greatest percent decrease, increase, and total volume will also be recorded
'this macro will automatically loop through each worksheet in a workbook

'declare variables
Dim i, j, LastRow, TickerTypes, MaxVolume As Long
Dim Ticker As String
Dim Volume, TickerOpen, TickerClose, change, ChangeMax, PercentMax As Double
Dim ws As Worksheet

'loops through all worksheets in the workbook and applies the ws prefix to all cell locations
For Each ws In ActiveWorkbook.Worksheets
With ws

'set last row in the active sheet
LastRow = .Cells(Rows.Count, "A").End(xlUp).Row

'start TickerTypes counter at 0
TickerTypes = 0

'Add headers for the new summery tables
.Cells(1, 9).Value = ("Ticker")
.Cells(1, 10).Value = ("Yearly Change")
.Cells(1, 11).Value = ("Percent Change")
.Cells(1, 12).Value = ("Total Stock Volume")
.Cells(1, 16).Value = ("Ticker")
.Cells(1, 17).Value = ("Value")


'Begin primary For loop skipping the header row
For i = 2 To LastRow
'error handling to deal with missing stock data
On Error Resume Next
'When ticker value matches ticker value in row below, adds up the volume
    If .Cells(i + 1, 1).Value = .Cells(i, 1) Then
        Volume = Volume + .Cells(i, 7)
        'Check for first row of the Ticker by checking Ticker value above, then sets the opening value
        If .Cells(i - 1, 1).Value <> .Cells(i, 1) Then
            TickerOpen = .Cells(i, 3)
        End If
'When ticker value is a mismatch to the value below it, indicating final row for the ticker
    ElseIf .Cells(i + 1, 1).Value <> .Cells(i, 1).Value Then
        'final volume added for the ticker
        Volume = Volume + .Cells(i, 7)
        'add 1 to the ticker types counter
        TickerTypes = 1 + TickerTypes
        'the close value is set for the ticker
        TickerClose = .Cells(i, 6)
        'calculates the yearly change and reports out to summery table
        .Cells(TickerTypes + 1, 10) = TickerClose - TickerOpen
        'sets percent cell to % formattings and reports out the percent change to summery table
        .Cells(TickerTypes + 1, 11).NumberFormat = "0.00%"
        .Cells(TickerTypes + 1, 11) = .Cells(TickerTypes + 1, 10) / TickerOpen
        'reports out the ticker symbol to the summery table
        .Cells(TickerTypes + 1, 9) = .Cells(i, 1).Value
        'reports out the final volume to the summery table
        .Cells(TickerTypes + 1, 12) = Volume
        'conditional formatting of the yearly change in the summery table
            If .Cells(TickerTypes + 1, 10) > 0 Then
                .Cells(TickerTypes + 1, 10).Interior.ColorIndex = 4
            ElseIf .Cells(TickerTypes + 1, 10) < 0 Then
                .Cells(TickerTypes + 1, 10).Interior.ColorIndex = 3
            End If
        'resets the ticker open, close, and volume values to zero in preparation for next row
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


End With
Next

End Sub
