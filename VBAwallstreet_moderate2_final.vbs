sub TheVBAofWallStreet()

    ' loop through each worksheet
    for each ws in Worksheets

    ' create the ticker and TSV columns and define them
    ws.range("I1").value = "Ticker"
    ws.range("J1").value = "Yearly Change"
    ws.range("K1").value = "Percent Change"
    ws.range("L1").value = "Total Stock Volume"
    dim Ticker as String
    dim YearlyChange as Double
    dim PercentChange as Double
    dim StockVolume as Double
    ' for the yearly change and percentages
    dim OpeningPrice as Double
    dim ClosingPrice as Double

    ' define the summary table, which will start from the second row
    dim Summary as Integer
    Summary = 2

    ' set stock volume to zero
    StockVolume = 0

    ' define the last row variable which will count the rows
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' loop through each row of the first column

        for i = 2 to LastRow

            ' define conditional for change in ticker

            ' if the two cells are not equal
            if ws.Cells(i + 1, 1).value <> ws.Cells(i, 1).value then
            ' assign the change as the new value of ticker
            Ticker = ws.cells(i, 1).value

            ' set the closing price and the opening price
            ClosingPrice = ws.cells(i, 6).value

            ' determine yearly change
            YearlyChange = ClosingPrice - OpeningPrice

            ' determine percent change
            PercentChange = YearlyChange / OpeningPrice

            ' assign corresponding stock volume to TSV
            StockVolume = StockVolume + ws.cells(i, 7).value

            ' print these to the summary table
            ws.range("I" & Summary).value = Ticker
            ws.range("L" & Summary).value = StockVolume
            WS.range("J" & Summary).value = YearlyChange
            ws.range("K" & Summary).value = PercentChange
            ws.range("K" & Summary).NumberFormat = "0.00%"
            'conditional formatting
                if ws.range("J" & Summary).value > 0 then
                ws.range("J" & Summary).Interior.ColorIndex = 4
                elseif ws.range("J" & Summary).value < 0 then
                ws.range("J" & Summary).Interior.ColorIndex = 3
                end if

            ' move to next row
            Summary = Summary + 1

            ' reset stock volume and opening price
            StockVolume = 0
            OpeningPrice = 0

            else
            StockVolume = StockVolume + ws.cells(i, 7).value
                if OpeningPrice = 0 then
                OpeningPrice = ws.cells(i, 3).value
                end if
            
            end if
        
        next i
    
    next ws

end sub