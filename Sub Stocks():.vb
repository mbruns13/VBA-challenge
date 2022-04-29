' need to fix: setting FirstOpen automatically from the start instead of manually starting it at the cell value (currently at line 25)
                'figure out how to loop through multiple sheets
                'figure out how to do max/min for bonus table

Sub Stocks():
    'set variables for data to go in summary table
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As LongLong
    YearlyChange = 0
    PercentChange = 0
    TotalStockVolume = 0
    
    'place headers in summary table columns
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    'fit column widths to fit new headers
    'NOTE ///// this will need to change when working with larger sheet
    Worksheets("A").Columns("I:L").AutoFit
    
    'set variables to calculate yearly and percent change
    Dim FirstOpen As Double
    Dim LastClose As Double
    FirstOpen = Cells(2, 3).Value

    'set summary table row as variable so output can move down a row as needed
    Dim SummaryRow As Integer
    SummaryRow = 2

    'set last row number as variable
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
 
    'loop through all stocks for one year and outputs:
    For i = 2 To LastRow
        'if the next row's ticker name does not match the current row's
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'set ticker variable to match that of the current row's
            Ticker = Cells(i, 1).Value
            'set last close variable to the value from current row's column 6
            LastClose = Cells(i, 6).Value
            'add the current row's stock volume to the count
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value
            'add current row's ticker name to the summary table
            Cells(SummaryRow, 9).Value = Ticker
            'calculate yearly change and add to summary table
            YearlyChange = LastClose - FirstOpen
            Cells(SummaryRow, 10).Value = YearlyChange
            'calculate percent change and add to summary table
            PercentChange = YearlyChange / FirstOpen
            Cells(SummaryRow, 11).Value = (PercentChange * 100) & "%"
            'add current row's total stock volume to the summary table
            Cells(SummaryRow, 12).Value = TotalStockVolume
            'move summary row variable down one row so it is ready for the next ticker info
            SummaryRow = SummaryRow + 1
            'reset stock volume so it's ready to begin counting for the next ticker
            TotalStockVolume = 0

            'set first open variable to the value from the NEXT row's column 3
            FirstOpen = Cells(i + 1, 3).Value

        Else
            'if current row and next row are from the same ticker, add current row's stock volume to running total
            TotalStockVolume = TotalStockVolume + Cells(i, 7).Value

        End If
    Next i
    
    'loop through summary table to color code yearly change column
    For i = 2 To LastRow
        'if cells contain positive value/increase
        If Cells(i, 10).Value > 0 Then
            'make cells green
            Cells(i, 10).Interior.ColorIndex = 4
        'if cells contain negative value/decrease
        ElseIf Cells(i, 10).Value < 0 Then
            'make cells red
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i




End Sub