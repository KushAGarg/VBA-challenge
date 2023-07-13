Attribute VB_Name = "Module1"
Sub stockanalysis():

    'Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets

        'Set variables for calculations
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim OpeningPrice As Double
        Dim ClosingPrice As Double
        Dim MaxPercentIncrease As Double
        Dim MaxPercentIncreaseTicker As String
        Dim MaxPercentDecrease As Double
        Dim MaxPercentDecreaseTicker As String
        Dim MaxTotalVolume As Double
        Dim MaxTotalVolumeTicker As String

        'Set location for variables
        Dim SummaryRow As Long
        SummaryRow = 2
        
        'Loop through all sheets to find last cell that is not empty
        Dim LastRow As Long
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Label columns and rows
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Do the calculations
        For i = 2 To LastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                OpeningPrice = ws.Cells(i, 3).Value
                ClosingPrice = ws.Cells(i, 6).Value
                YearlyChange = ClosingPrice - OpeningPrice
                If OpeningPrice <> 0 Then
                    PercentChange = (YearlyChange / OpeningPrice) * 100
                Else
                    PercentChange = 0
                End If
        
                'Print results in labeled columns
                ws.Cells(SummaryRow, 9).Value = Ticker
                ws.Cells(SummaryRow, 10).Value = YearlyChange
                ws.Cells(SummaryRow, 11).Value = PercentChange
                ws.Cells(SummaryRow, 12).Value = TotalVolume
                
                'Format number style in Percent Change column
                ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
        
                'Color fill Yearly Change column: Red for negative and green for positive
                If YearlyChange > 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 3
                ElseIf YearlyChange < 0 Then
                    ws.Cells(SummaryRow, 10).Interior.ColorIndex = 4
                End If
        
                'Add 1 to the summary row count
                SummaryRow = SummaryRow + 1
                
                'Reset value
                TotalVolume = 0
            
            'Else if in next ticker name, enter new ticker stock volume
            Else
                TotalVolume = TotalVolume + Cells(i, 7).Value
            End If
            
        Next i

        'Find the stock with the greatest percent increase
        MaxPercentIncrease = Application.WorksheetFunction.Max(ws.Range("K2:K" & SummaryRow))
        MaxPercentIncreaseTicker = ws.Cells(Application.WorksheetFunction.Match(MaxPercentIncrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
        ws.Cells(2, 16).Value = MaxPercentIncreaseTicker
        ws.Cells(2, 17).Value = MaxPercentIncrease
        ws.Cells(2, 17).NumberFormat = "0.00%"
    
        'Find the stock with the greatest percent decrease
        MaxPercentDecrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & SummaryRow))
        MaxPercentDecreaseTicker = ws.Cells(Application.WorksheetFunction.Match(MaxPercentDecrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
        ws.Cells(3, 16).Value = MaxPercentDecreaseTicker
        ws.Cells(3, 17).Value = MaxPercentDecrease
        ws.Cells(3, 17).NumberFormat = "0.00%"
    
        'Find the stock with the greatest total volume
        MaxTotalVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & SummaryRow))
        MaxTotalVolumeTicker = ws.Cells(Application.WorksheetFunction.Match(MaxTotalVolume, ws.Range("L2:L" & SummaryRow), 0) + 1, 9).Value
        ws.Cells(4, 16).Value = MaxTotalVolumeTicker
        ws.Cells(4, 17).Value = MaxTotalVolume
    
    Next ws
    
End Sub
