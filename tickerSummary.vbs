Sub tickerSummary()
    
    For Each ws In Worksheets
        Dim current_Summary_Row As Long
        current_Summary_Row = 2
        
        'Set the output headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'find size of the worksheet
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        'BONUS: Set worksheet level comparison values, set comparisons to zero to avoid null comparison errors
        Dim GreatIncreaseTicker As String
        Dim GreatIncreaseValue As Double
        GreatIncreaseValue = 0
        Dim GreatDecreaseTicker As String
        Dim GreatDecreaseValue As Double
        GreatDecreaseValue = 0
        Dim GreatTotalVolumeTicker As String
        Dim GreatTotalVolume As Double
        GreatTotalVolume = 0
        
        'find all tickers in worksheet and store into Array
        Dim tickerArray() As String
        Dim currTicker As String
        Dim tickerCheck As Long
        Dim tickerArrayIndex As Long
        tickerArrayIndex = 0
        'set the first ticker to avoid null errors
        ReDim tickerArray(tickerArrayIndex + 1)
        tickerArray(0) = ws.Cells(2, 1)
        For i = 3 To LastRow
            tickerCheck = 0
            Dim tickVal As String
            tickVal = ws.Cells(i, 1).Value
            For Each ticker In tickerArray
                If ticker = tickVal Then
                    tickerCheck = 1
                End If
            Next ticker
            If tickerCheck = 0 Then
                ReDim Preserve tickerArray(tickerArrayIndex + 1)
                tickerArray(tickerArrayIndex + 1) = tickVal
                tickerArrayIndex = tickerArrayIndex + 1
            End If
        Next i
        
        'find highest and lowest year dates and record their row
        For Each ticker In tickerArray
            Dim lowDate As Long
            Dim lowDateRow
            Dim highDate As Long
            Dim highDateRow
            Dim volume As Double
            lowDate = 0
            lowDateRow = 0
            highDate = 0
            highDateRow = 0
            volume = 0
            currTicker = ticker
            For i = 2 To LastRow
                If ws.Cells(i, 1).Value = currTicker Then
                    volume = volume + ws.Cells(i, 7)
                    Dim test As Long
                    test = ws.Cells(i, 2).Value
                    If lowDate = 0 Or ws.Cells(i, 2).Value < lowDate Then
                        lowDate = ws.Cells(i, 2).Value
                        lowDateRow = i
                    End If
                    If highDate = 0 Or ws.Cells(i, 2).Value > highDate Then
                        highDate = ws.Cells(i, 2).Value
                        highDateRow = i
                    End If
                End If
            Next i
            
            'Add findings to summary sheet
            'Run the comparisons from the first date to last date in range
            Dim yearChange As Double
            yearChange = ws.Cells(highDateRow, 6).Value - ws.Cells(lowDateRow, 3).Value
            ws.Cells(current_Summary_Row, 9).Value = currTicker
            
            ws.Cells(current_Summary_Row, 10).Value = yearChange
            'Format year change based on positive value
            If yearChange > 0 Then
                ws.Cells(current_Summary_Row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(current_Summary_Row, 10).Interior.ColorIndex = 3
            End If
            
            If ws.Cells(lowDateRow, 3).Value <> 0 Then
                ws.Cells(current_Summary_Row, 11).Value = yearChange / ws.Cells(lowDateRow, 3).Value
                'BONUS: Compare to Greatest Increase and Decrease
                If (yearChange / ws.Cells(lowDateRow, 3).Value) > GreatIncreaseValue Then
                    GreatIncreaseTicker = currTicker
                    GreatIncreaseValue = yearChange / ws.Cells(lowDateRow, 3).Value
                End If
                If (yearChange / ws.Cells(lowDateRow, 3).Value) < GreatDecreaseValue Then
                    GreatDecreaseTicker = currTicker
                    GreatDecreaseValue = yearChange / ws.Cells(lowDateRow, 3).Value
                End If
            Else
                'If we have invalid data where a stock opened its year with $0 value mark as such!
                ws.Cells(current_Summary_Row, 11).Value = "ERR - No Opening Value"
            End If
            
            ws.Cells(current_Summary_Row, 11).NumberFormat = "0.00%"
            ws.Cells(current_Summary_Row, 12).Value = volume
            If volume > GreatTotalVolume Then
                GreatTotalVolumeTicker = currTicker
                GreatTotalVolume = volume
            End If
            
            current_Summary_Row = current_Summary_Row + 1
        Next ticker
        
        'Write Greatest Values
        'Write Column and Row Headers
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        'Write greatest values found
        ws.Cells(2, 16).Value = GreatIncreaseTicker
        ws.Cells(3, 16).Value = GreatDecreaseTicker
        ws.Cells(4, 16).Value = GreatTotalVolumeTicker
        ws.Cells(2, 17).Value = GreatIncreaseValue
        ws.Cells(3, 17).Value = GreatDecreaseValue
        ws.Cells(4, 17).Value = GreatTotalVolume
        
    Next ws
        
End Sub

