Sub multiple_year_stock_data()

For Each ws In Worksheets

        Dim max As Double
            max = 0
        Dim maxTicker As String
        Dim min As Double
            min = 0
        Dim minTicker As String
        Dim totalMax As Double
            totalMax = 0
        Dim totalTicker As String
        Dim stockName As String
        Dim stockStart As Double
        Dim stockEnd As Double
            stockEnd = 0
        Dim stockChange As Double
            stockChange = 0
        Dim stockVol As Double
            stockVol = 0
        Dim stockPercentage As Double
            stockPercentage = 0
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
         For i = 2 To LastRow
        
        
            
            Dim comparisionChecker As Boolean
        
            If comparisionChecker = False Then
                
                stockStart = Cells(i, 3).Value
                
                comparisionChecker = True
                
            End If
            
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                
                stockName = Cells(i, 1).Value
                stockVol = stockVol + Cells(i, 7).Value
                stockEnd = Cells(i, 6).Value
                stockChange = stockEnd - stockStart
                stockPercentage = stockChange / stockStart
                Range("I" & Summary_Table_Row).Value = stockName
                Range("J" & Summary_Table_Row).Value = stockChange
                Range("K" & Summary_Table_Row).Value = stockPercentage
                Range("L" & Summary_Table_Row).Value = stockVol
                Summary_Table_Row = Summary_Table_Row + 1
                stockVol = 0
                stockPercentage = 0
                stockChange = 0
                comparisionChecker = False
            
                
            Else
                stockVol = stockVol + Cells(i, 7).Value
                stockChange = stockEnd - stockStart
                
            End If
            
        
            Next i

                    For j = 2 To LastRow

                        If Cells(j, 10).Value > 0 Then
                        Cells(j, 10).Interior.ColorIndex = 4
                        ElseIf Cells(j, 10).Value < 0 Then
                        Cells(j, 10).Interior.ColorIndex = 3
                        ElseIf Cells(j, 10).Value = 0 Then
                        Cells(j, 10).Interior.ColorIndex = 2
                    End If
                    Next j


                    For k = 2 To LastRow
                    
                        If ws.Cells(k, 11) > max Then
                        max = ws.Cells(k, 11)
                        maxTicker = ws.Cells(k, 9)
                    End If
                    Next k
    
                    For l = 2 To LastRow
                    
                        If ws.Cells(l, 11) < min Then
                        min = ws.Cells(l, 11)
                        minTicker = ws.Cells(l, 9)
                    End If
                    Next l

                    For m = 2 To LastRow
                        If ws.Cells(m, 12) > totalMax Then
                        totalMax = ws.Cells(m, 12)
                        totalTicker = ws.Cells(m, 9)
                    End If
                    Next m

                                            ws.Cells(2, 16) = maxTicker
                                            ws.Cells(3, 16) = minTicker
                                            ws.Cells(4, 16) = totalTicker
                                            ws.Cells(2, 17).Value = max
                                            ws.Cells(3, 17).Value = min
                                            ws.Cells(4, 17).Value = totalMax




Next ws

End Sub
