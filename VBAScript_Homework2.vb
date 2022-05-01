Sub Stocks():
    'set variables for data to go in summary table
    Dim Ticker As String
    Dim YearlyChange, PercentChange As Double
    Dim TotalStockVolume As LongLong
    
    'set variables for bonus ("Greatest % increase", "Greatest % decrease", and "Greatest total volume")
    Dim GreatestInc, GreatestDec As Double
    Dim IncTicker, DecTicker, GreatestTotalTicker As String
    Dim GreatestTotal As LongLong
    Dim MatchRow As Integer
    
    Dim ws As Worksheet
    
    'start loop through all worksheets
    For Each ws In Worksheets
    
        'place headers in summary table columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        Worksheets(ws.Name).Columns("J:L").AutoFit
    
        'set variables to calculate yearly and percent change
        Dim FirstOpen, LastClose As Double
        FirstOpen = ws.Cells(2, 3).Value

        'set summary table row as variable so output can move down a row as needed
        Dim SummaryRow As Integer
        SummaryRow = 2
 
        'set last row number as variable
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
 
        'loop through all stocks for one year and output:
        For i = 2 To LastRow
            'if the next row's ticker name does not match the current row's
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'set ticker variable to match that of the current row's
                Ticker = ws.Cells(i, 1).Value
                'set last close variable to the value from current row's column 6
                LastClose = ws.Cells(i, 6).Value
                'add the current row's stock volume to the count
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
                'add current row's ticker name to the summary table
                ws.Cells(SummaryRow, 9).Value = Ticker
                'calculate yearly change and add to summary table
                YearlyChange = LastClose - FirstOpen
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                'calculate percent change and add to summary table and format
                PercentChange = YearlyChange / FirstOpen
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 11).Value = FormatPercent(ws.Cells(SummaryRow, 11))
                'add current row's total stock volume to the summary table
                ws.Cells(SummaryRow, 12).Value = TotalStockVolume
                'move summary row variable down one row so it is ready for the next ticker info
                SummaryRow = SummaryRow + 1
                'reset stock volume so it's ready to begin counting for the next ticker
                TotalStockVolume = 0

                'set first open variable to the value from the NEXT row's column 3
                FirstOpen = ws.Cells(i + 1, 3).Value

            Else
                'if current row and next row are from the same ticker, add current row's stock volume to running total
                TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value

            End If
        Next i
    
        'loop through summary table to color code yearly change column
        For i = 2 To LastRow
            'if cells contain positive value/increase
            If ws.Cells(i, 10).Value > 0 Then
                'make cells green
                ws.Cells(i, 10).Interior.ColorIndex = 4
            'if cells contain negative value/decrease
            ElseIf ws.Cells(i, 10).Value < 0 Then
                'make cells red
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
    Next ws

    For Each ws In Worksheets
        'place headers for bonus section
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        Worksheets(ws.Name).Columns("O").AutoFit
        'loop through all sheets to find greatest increase/decrease/total stock volume
        
        'find greatest % increase, add to cell Q2, & format
        GreatestInc = WorksheetFunction.Max(ws.Range("K2:K" & LastRow))
        ws.Range("Q2").Value = GreatestInc
        ws.Range("Q2").Value = FormatPercent(ws.Range("Q2"))
        'match GreatestInc with its ticker value, add to cell P2
        MatchRow = (WorksheetFunction.Match(GreatestInc, ws.Range("K2:K" & LastRow), 0) + 1)
        IncTicker = ws.Cells(MatchRow, 9).Value
        ws.Range("P2").Value = IncTicker

        'find greatest % decrease, add to cell Q3, & format
        GreatestDec = WorksheetFunction.Min(ws.Range("K2:K" & LastRow))
        ws.Range("Q3").Value = GreatestDec
        ws.Range("Q3").Value = FormatPercent(ws.Range("Q3"))
        'match GreatestDec with its ticker value, add to cell P3
        MatchRow = (WorksheetFunction.Match(GreatestDec, ws.Range("K2:K" & LastRow), 0) + 1)
        DecTicker = ws.Cells(MatchRow, 9).Value
        ws.Range("P3").Value = DecTicker

        'find greatest total stock volume, add to cell Q4
        GreatestTotal = WorksheetFunction.Max(ws.Range("L2:L" & LastRow))
        ws.Range("Q4").Value = GreatestTotal
        'match GreatestTotal with its ticker value, add to cell P4
        MatchRow = (WorksheetFunction.Match(GreatestTotal, ws.Range("L2:L" & LastRow), 0) + 1)
        GreatestTotalTicker = ws.Cells(MatchRow, 9).Value
        ws.Range("P4").Value = GreatestTotalTicker
    
    Next ws

End Sub