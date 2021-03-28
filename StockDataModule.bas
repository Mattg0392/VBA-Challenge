Attribute VB_Name = "Module1"
Sub StockData():
    
    For Each ws In Worksheets
        
        Dim startValue As Double
        
        Dim isStartValue As Boolean
        
        isStartValue = True
        
        Dim lastRow As Double
        
        Dim yearlyChange As Double
        
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim tickerSymbol As String
        
        Dim tickerTotal As Double
        
        Dim summaryRow As Integer
        summaryRow = 2
        
        For Row = 2 To lastRow
        
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
            
                tickerSymbol = ws.Cells(Row, 1).Value
            
                tickerTotal = tickerTotal + ws.Cells(Row, 7).Value
            
                ws.Range("I" & summaryRow).Value = tickerSymbol
            
                ws.Range("L" & summaryRow).Value = tickerTotal
                
                yearlyChange = ws.Cells(Row, 6).Value - startValue
                
                ws.Range("J" & summaryRow).Value = yearlyChange
                
                If yearlyChange = 0 Then
                    
                    ws.Range("K" & summaryRow).Value = 0
                    
                Else
                
                    ws.Range("K" & summaryRow).Value = (yearlyChange / startValue) * 100
                    
                End If
            
                summaryRow = summaryRow + 1
            
                tickerTotal = 0
                
                isStartValue = True
            
            Else
            
                If isStartValue And ws.Cells(Row, 3).Value <> 0 Then
                    
                    startValue = ws.Cells(Row, 3).Value
                    
                    isStartValue = False
                    
                End If
                
                tickerTotal = tickerTotal + ws.Cells(Row, 7).Value
                
                
            
            End If
        
        Next Row
    
    Next ws
    
End Sub

