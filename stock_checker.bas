Attribute VB_Name = "Module1"
Sub stock_checker()
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim Ticker As String
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim QuarterlyChange As Double
    Dim PercentChange As Double
    Dim TotalStockVolume As Double
    Dim j As Long
    Dim GreatestPerInc As Double
    Dim GreatestPerDec As Double
    Dim GreatestTotVal As Double
    Dim GreatestPerIncTicker As String
    Dim GreatestPerDecTicker As String
    Dim GreatestTotValTicker As String
    
    For Each ws In ThisWorkbook.Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Set headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        TotalStockVolume = 0
        j = 2
        
        For i = 2 To LastRow
            Ticker = ws.Cells(i, 1).Value
            
            'check for new ticker
            If Ticker <> ws.Cells(i - 1, 1).Value Then
                OpeningPrice = ws.Cells(i, 3).Value
            End If
            
            'accumulate total stock volume
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
            'check end of ticker
            If Ticker <> ws.Cells(i + 1, 1).Value Then
                ClosingPrice = ws.Cells(i, 6).Value
                QuarterlyChange = ClosingPrice - OpeningPrice
                PercentChange = ((ClosingPrice - OpeningPrice) / OpeningPrice)
                
                ws.Cells(j, 9).Value = Ticker
                ws.Cells(j, 10).Value = QuarterlyChange
                ws.Cells(j, 11).Value = PercentChange
                ws.Cells(j, 12).Value = TotalStockVolume
                
                TotalStockVolume = 0
                
                If QuarterlyChange > 0 Then
                    ws.Cells(j, 10).Interior.Color = vbGreen
                
                ElseIf QuarterlyChange < 0 Then
                    ws.Cells(j, 10).Interior.Color = vbRed
                    
                ElseIf QuarterlyChange = 0 Then
                    ws.Cells(j, 10).Interior.Color = xlNone
                
                End If
                j = j + 1
            End If
        Next i
        
        LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'set values
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        GreatestPerInc = 0
        GreatestPerDec = 0
        GreatestTotVal = 0
        
        For i = 2 To LastRow
        
            If ws.Cells(i, 11).Value > GreatestPerInc Then
                GreatestPerInc = ws.Cells(i, 11).Value
                GreatestPerIncTicker = ws.Cells(i, 9).Value
            
            ElseIf ws.Cells(i, 11).Value < GreatestPerDec Then
                GreatestPerDec = ws.Cells(i, 11).Value
                GreatestPerDecTicker = ws.Cells(i, 9).Value
                
            End If
            
            If ws.Cells(i, 12).Value > GreatestTotVal Then 
                GreatestTotVal = ws.Cells(i, 12).Value
                GreatestTotValTicker = ws.Cells(i, 9).Value
                
            End If
        Next i
        ws.Cells(2, 16) = GreatestPerIncTicker
        ws.Cells(2, 17) = GreatestPerInc
        ws.Cells(3, 16) = GreatestPerDecTicker
        ws.Cells(3, 17) = GreatestPerDec
        ws.Cells(4, 16) = GreatestTotValTicker
        ws.Cells(4, 17) = GreatestTotVal
    Next ws
End Sub