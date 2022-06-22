Attribute VB_Name = "Module1"
Sub StockMarketAnalyst()

For Each ws In Worksheets

'Adding and Formatting The headers to the Summary Table
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"
ws.Range("J:M").Columns.AutoFit

'Defining some variables
TotalStockVolume = 0
SummaryTableRow = 2
Ticker = ""
YearlyChange = 0

'Figuring out last row
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'Loops through the rows using column 1
For r = 2 To LastRow

'If Ticker Value Changes
If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
    Ticker = ws.Cells(r, 1).Value
    Match = WorksheetFunction.Match(Ticker, ws.Range("A:A"), 0) 'will give us the first row where each ticker symbol is found
    YearlyChange = ws.Cells(r, 6).Value - ws.Cells(Match, 3)
    PercentChange = (YearlyChange) / (ws.Cells(Match, 3))
    TotalStockVolume = TotalStockVolume + ws.Cells(r, 7).Value
  
    
    ws.Range("J" & SummaryTableRow) = Ticker
    ws.Range("K" & SummaryTableRow) = YearlyChange
            If ws.Range("K" & SummaryTableRow) < 0 Then
            ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 3
            ElseIf ws.Range("K" & SummaryTableRow) > 0 Then
            ws.Range("K" & SummaryTableRow).Interior.ColorIndex = 4
            End If 'setting coloring format
            ws.Range("K:K").NumberFormat = "0.00"
    ws.Range("L" & SummaryTableRow) = PercentChange
             ws.Range("L:L").NumberFormat = "0.00%"  'setting percentage format
    ws.Range("M" & SummaryTableRow) = TotalStockVolume
    
    
    YearlyChange = 0 'resetting the yearly change to 0
    PercentChange = 0 'resetting the percent change to 0
    TotalStockVolume = 0 'resetting the total stock volume to 0
    SummaryTableRow = SummaryTableRow + 1 'starting on the next row for the next loop

'If Ticker Value Doesn't Change
Else
    TotalStockVolume = TotalStockVolume + ws.Cells(r, 7).Value
End If

Next r

''''BONUS SECTION

'Adding and formatting the headers
ws.Range("Q1").Value = "Ticker"
ws.Range("R1").Value = "Value"
ws.Range("P2").Value = "Greatest % Increase"
ws.Range("P3").Value = "Greatest % Decrease"
ws.Range("P4").Value = "Greatest Total Volume"

'Finding Max Values
GreatestIncrease = WorksheetFunction.Max(ws.Range("L:L"))
GreatestDecrease = WorksheetFunction.Min(ws.Range("L:L"))
GreatestTotalVolume = WorksheetFunction.Max(ws.Range("M:M"))

'Placing Max Values
ws.Range("R2").Value = GreatestIncrease
    ws.Range("R2").NumberFormat = "0.00%"
ws.Range("R3").Value = GreatestDecrease
    ws.Range("R3").NumberFormat = "0.00%"
ws.Range("R4").Value = GreatestTotalVolume

'Autofit Columns
ws.Range("P:R").Columns.AutoFit

'Find Ticker Value Using Match Function
Match = WorksheetFunction.Match(ws.Range("R2").Value, ws.Range("L:L"), 0)
ws.Range("Q2").Value = Range("J" & Match).Value

Match = WorksheetFunction.Match(ws.Range("R3").Value, ws.Range("L:L"), 0)
ws.Range("Q3").Value = Range("J" & Match).Value

Match = WorksheetFunction.Match(ws.Range("R4").Value, ws.Range("M:M"), 0)
ws.Range("Q4").Value = Range("J" & Match).Value

Next ws

End Sub

