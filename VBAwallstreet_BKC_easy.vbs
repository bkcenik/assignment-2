sub TheVBAofWallStreet()

    ' loop through each worksheet
    for each ws in Worksheets

    ' create the ticker and TSV columns and define them
    ws.range("I1").value = "Ticker"
    ws.range("J1").value = "Total Stock Volume"
    dim Ticker as String
    dim StockVolume as Double

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

            ' assign corresponding stock volume to TSV
            StockVolume = ws.cells(i, 7).value

            ' print these to the summary table
            ws.range("I" & Summary).value = Ticker
            ws.range("J" & Summary).value = StockVolume

            ' move to next row
            Summary = Summary + 1
            
            ' reset stock volume
            StockVolume = 0

            ' if ticker doesn't change, continue adding up the stock volume
            else
            StockVolume = StockVolume + ws.cells(i, 7).value
            
            end if
        
        next i
    
    next ws

end sub