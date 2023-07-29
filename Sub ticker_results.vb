Sub ticker_results()
    For Each ws In Worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
     Dim currentName, nextName, previousName As String
     Dim openamount, closeamount, greatestIncrease, greatestDecrease As Integer
     Dim currentStockVol, currentStockVolTotal, greatestVolume As Double
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
     stockname = 2
     yearlyChange = 2
     currentStockVol = 0
     currentStockVolTotal = 0
    For row_num = 2 To lastrow
        currentName = ws.Cells(row_num, 1).Value
        nextName = ws.Cells(row_num + 1, 1).Value
        previousName = ws.Cells(row_num - 1, 1).Value
        currentStockVol = ws.Cells(row_num, 7).Value
        currentStockVolTotal = currentStockVolTotal + currentStockVol
        stockStart = ws.Cells(row_num, 3).Value
        stockClose = ws.Cells(row_num, 6).Value
            
            If (currentName <> previousName And currentName = nextName) Then
                ws.Range("I" & stockname).Value = currentName
                stockStart_figure = stockStart
                End If
            
                    If (currentName = previousName And currentName <> nextName) Then
                    stockClose_figure = stockClose
                    ws.Range("L" & stockname).Value = currentStockVolTotal
                    stockname = stockname + 1
                    currentStockVolTotal = 0
                    ws.Range("J" & yearlyChange).Value = stockClose_figure - stockStart_figure
                    ws.Range("K" & yearlyChange).Value = ws.Range("J" & yearlyChange).Value / stockStart_figure
                    ws.Range("K" & yearlyChange).NumberFormat = "0.00%"
                    ws.Range("L" & yearlyChange).NumberFormat = "#,##0"
                    yearlyChange = yearlyChange + 1
                    End If
    Next row_num
    
    newlastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For newrow_num = 2 To newlastrow
                    If (ws.Cells(newrow_num, 10) >= 0) Then
                        ws.Cells(newrow_num, 10).Interior.ColorIndex = 4
                    End If
                    
                    If (ws.Cells(newrow_num, 10) < 0) Then
                        ws.Cells(newrow_num, 10).Interior.ColorIndex = 3
                    End If
                    
                ws.Cells(2, 16).Value = WorksheetFunction.Max(ws.Range("K2" & ":" & "K" & newlastrow))
                ws.Cells(3, 16).Value = WorksheetFunction.Min(ws.Range("K2" & ":" & "K" & newlastrow))
                ws.Cells(4, 16).Value = WorksheetFunction.Max(ws.Range("L2" & ":" & "L" & newlastrow))
                    
                If (ws.Cells(2, 16).Value = ws.Cells(newrow_num, 11)) Then
                ws.Cells(2, 15).Value = ws.Cells(newrow_num, 9)
                End If
                
                If (ws.Cells(3, 16).Value = ws.Cells(newrow_num, 11)) Then
                ws.Cells(3, 15).Value = ws.Cells(newrow_num, 9)
                End If
                
                If (ws.Cells(4, 16).Value = ws.Cells(newrow_num, 12)) Then
                ws.Cells(4, 15).Value = ws.Cells(newrow_num, 9)
                End If
    Next newrow_num
        
        
        ws.Range("P2", "P3").NumberFormat = "0.00%"
        ws.Cells(4, 16).NumberFormat = "#,##0"
        ws.Columns("A:P").AutoFit
        
    Next ws
End Sub

