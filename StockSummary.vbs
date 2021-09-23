Sub SheetSummary()

For Each ws In Worksheets 'For all Sheets in the Worksheet
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row 'Lastrow with data in the sheet
    Stockvolume = 0 'Variable to keep track of stock volume per Sticker
    SummaryRow = 1 'Row where we want to put the Summary row.1st row will be Header
    For Row = 2 To lastrow
        'If Current cell value Matches the next cell value in column i.e. same ticker symbol
        If ws.Cells(Row, 1).Value = ws.Cells(Row + 1, 1) Then
            Stockvolume = Stockvolume + ws.Cells(Row, 7).Value 'Sumup the Stock volume
                If Row = 2 Then 'For First Sticker 2nd row i.e 1st (opening) date
                    OpeningValue = ws.Cells(Row, 3).Value 'Store the Opening Value
                End If
        Else 'If Current cell value does not matchthe next cell value in column i.e. ticker symbol Changed
            Stockvolume = Stockvolume + ws.Cells(Row, 7).Value 'Sumup the stock volume for the last row of the ticker
            ClosingValue = ws.Cells(Row, 3).Value ' Get Closing Value from the last row of the sticker
            If SummaryRow = 1 Then 'For the first time in the sheet when we are summariazing,add Headers for Summary
                ws.Cells(SummaryRow, 9) = "Ticker"
                ws.Cells(SummaryRow, 10) = "Yearly Change"
                ws.Cells(SummaryRow, 11) = "Percent Change"
                ws.Cells(SummaryRow, 12) = "Total Stock Volume"
            End If
            SummaryRow = SummaryRow + 1 'Increase the value of SummaryRow counter, so that me move to the next row
            'Put the Sticker Value, for the Current Sticker in the Summary Table
            ws.Cells(SummaryRow, 9).Value = ws.Cells(Row, 1).Value
            ' Calculate Yearly Change for the Ticker and put it in the summary Table
            ws.Cells(SummaryRow, 10).Value = ClosingValue - OpeningValue
            If (ClosingValue - OpeningValue) < 0 Then ' If Yearly Change is negative
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3 ' Format cell to Red color
            Else ' If Yearly Change is not negative
                ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4 ' Format cell to Green
            End If
                
            'Calculate Percent Change
            If OpeningValue = 0 Then
                ws.Cells(SummaryRow, 11).Value = 0
                'If opening Value is 0,We cannot caluclate % Change, hence default to 0
            Else 'Calculate and put % Change in Summary Table
                ws.Cells(SummaryRow, 11).Value = (ClosingValue - OpeningValue) / OpeningValue
            End If
            ws.Cells(SummaryRow, 11).NumberFormat = "0.00%" ' Format % Change as Percent
            ws.Cells(SummaryRow, 12).Value = Stockvolume 'Put Total Stock Volume in Summary Table
            Stockvolume = 0 ' Reset StockVolume to 0 for next Sticker
            OpeningValue = ws.Cells(Row + 1, 3) ' Set OpeningValue for NextSticker
          End If
     
    Next Row ' Go to Nest row in the sheet
            
'Bonus Task

' Activate Current Sheet, so that all Ranges refrenced in Application.WorksheetFunction apply to current sheet
ws.Activate

GreatestIncrease = ws.Application.WorksheetFunction.Max(Columns("K")) 'Get Greatest % Increase from Summary Table
GreatestDecrease = ws.Application.WorksheetFunction.Min(Columns("K")) 'Get Greatest % Decrease from Summary Table
GreatestVolume = ws.Application.WorksheetFunction.Max(Columns("L")) 'Get Greatest Total Volume from Summary Table

'Put Header Values for the Bonus Summary Table
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'Find the row with Greatest % Increase from Summary Table
GreatestIncreaseRow = ws.Application.WorksheetFunction.Match(GreatestIncrease, Columns("K"), 0)
'Gets the StickerValue from the row that contains Greatest % Increase from Summary Table
'Puts it in the Bonus Summary Table
ws.Cells(2, 16).Value = ws.Cells(GreatestIncreaseRow, 9)
'Puts Greatest % Increase value in the Bonus Summary Table
ws.Cells(2, 17).Value = GreatestIncrease
ws.Cells(2, 17).NumberFormat = "0.00%" ' Formats to Percent

'Find the row with Greatest % Decrease from Summary Table
GreatestDecreaseRow = ws.Application.WorksheetFunction.Match(GreatestDecrease, Columns("K"), 0)
'Gets the StickerValue from the row that contains Greatest % Decrease from Summary Table
'Puts it in the Bonus Summary Table
ws.Cells(3, 16).Value = ws.Cells(GreatestDecreaseRow, 9)
'Puts Greatest % Increase value in the Bonus Summary Table
ws.Cells(3, 17).Value = GreatestDecrease
ws.Cells(3, 17).NumberFormat = "0.00%" ' Formats to Percent

'Find the row with Greatest TotalVolume from Summary Table
GreatestVolumeRow = ws.Application.WorksheetFunction.Match(GreatestVolume, Columns("L"), 0)
'Gets the StickerValue from the row that contains Greatest Total Volume from Summary Table
'Puts it in the Bonus Summary Table
ws.Cells(4, 16).Value = ws.Cells(GreatestVolumeRow, 9)
'Puts Greatest Total Volume value in the Bonus Summary Table
ws.Cells(4, 17).Value = GreatestVolume

Next ws ' Go to Next Worksheet

End Sub




