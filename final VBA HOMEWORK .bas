Attribute VB_Name = "Module1"
   
   Sub VbaChallenge()

For Each ws In Worksheets
        
    Dim Ticker As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim SummaryRow As Long
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
                                
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
            
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    OpeningPrice = ws.Cells(2, 3).Value
    
    SummaryRow = 2
     
    For i = 2 To LastRow
             
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
     
            Ticker = ws.Cells(i, 1).Value
            ' OpeningPrice = ws.Cells(i, 3).Value
            ClosingPrice = ws.Cells(i, 6).Value
            YearlyChange = ClosingPrice - OpeningPrice
            
           
            If OpeningPrice <> 0 Then
                PercentChange = (YearlyChange / OpeningPrice)
            Else
                PercentChange = 0
            End If
                
            ws.Cells(SummaryRow, 9).Value = Ticker
            ws.Cells(SummaryRow, 10).Value = YearlyChange
            ws.Cells(SummaryRow, 11).Value = PercentChange
            'ws.Cells(SummaryRow, 12).Value = Application.WorksheetFunction.Sum (ws.Range(ws.Cells(i - Volume + 1, 7), wsCells(i, 7)))
            ws.Cells(SummaryRow, 11).NumberFormat = "0.00%"
            
            
            
            If YearlyChange > 0 Then
                ws.Cells(SummaryRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
            
            
            
            SummaryRow = SummaryRow + 1
            TotalVolume = 0
                 
        Else
                 
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
        End If
                 
    Next i
                    
    ' Find the maximum percent increase
    Dim MaxPercentIncrease As Double
    Dim MaxPercentTicker As String
    
    MaxPercentIncrease = WorksheetFunction.Max(ws.Range("K2:k" & SummaryRow))
    MaxPercentTicker = ws.Cells(WorksheetFunction.Match(MaxPercentIncrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
    
    ws.Cells(2, 16).Value = MaxPercentTicker
    ws.Cells(2, 17).Value = MaxPercentIncrease
    ws.Cells(2, 17).NumberFormat = "0.00%"
    
    'Find the maxium percent decrease
    Dim MaxPercentDecrease As Double
    Dim MaxPercentDecreaseTicker As String
    
    MaxPercentDecrease = WorksheetFunction.Min(ws.Range("k2:k" & SummaryRow))
    MaxPercentDecreaseTicker = ws.Cells(WorksheetFunction.Match(MaxPercentDecrease, ws.Range("K2:K" & SummaryRow), 0) + 1, 9).Value
    ws.Cells(3, 16).Value = MaxPercentDecreaseTicker
    ws.Cells(3, 17).Value = MaxPercentDecrease
    ws.Cells(3, 17).NumberFormat = "0.00%"
    
    ' Find the stock with the greatest total volume
    Dim MaxTotalVolume As Double
    Dim MaxTotalVolumeTicker As String
    
    MaxTotalVolume = WorksheetFunction.Max(ws.Range("L2:L" & SummaryRow))
    ' MaxTotalVolumeTicker = ws.Cells(WorksheetFunction.Match(MaxTotalVolume, ws.Range("L2:L" & SummaryRow), 0) + 1, 9).Value
    ws.Cells(4, 17).Value = MaxTotalVolume
    
    
    
    MsgBox ("Got here")
                

Next ws
                
End Sub

                
                
                
                
                
                
                
                
                
                
                
                
                
                 
             
            
             
             

    
    

    
    

