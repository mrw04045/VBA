Sub StockData()

    For Each ws In Worksheets

        'Create headers Ticker, Total Stock Volume
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Create variables
            Dim Ticker As String
            Dim TotalVolume As Variant
            Dim row As Integer
            Dim OpeningPrice As Double
            Dim ClosingPrice As Double
            Dim YearlyChange As Double
            Dim PercentChange As Double
            
        'Set initial total stock volume to zero
        TotalVolume = 0
        
        'Set intial opening price
        OpeningPrice = ws.Range("C2").Value
        
        'Set initial row value equal to 2
        row = 2
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Loop through Rows
        For i = 2 To LastRow
            
            'Define variable ticker
            Ticker = ws.Cells(i, 1).Value
            
            'Calculate total stock volume
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'If opening price is not zero, then go continue with following nested If statement
            If ws.Cells(i, 3).Value <> 0 Then
            
               'If following ticker symbol doesn't match current symbol
               If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
                   'Print ticker symbol under ticker header
                   ws.Cells(row, 9).Value = Ticker
                   
                   'Print total volume under total stock volume header
                   ws.Cells(row, 12).Value = TotalVolume
                   
                   'Reset total stock volume to 0
                   TotalVolume = 0
                   
                   'Define closing price
                   ClosingPrice = ws.Cells(i, 6).Value
                   
                   'Define yearly change
                   YearlyChange = ClosingPrice - OpeningPrice
                   
                   'Print yearly change under yearly change header
                   ws.Cells(row, 10).Value = YearlyChange
                   
                       'If yearly change is positive (>0)
                       If YearlyChange > 0 Then
                       
                       'Then Fill cell with green
                       ws.Cells(row, 10).Interior.ColorIndex = 4
                       
                       'Otherwise fill it with red to show negative change
                       Else
                       ws.Cells(row, 10).Interior.ColorIndex = 3
                       
                       End If
                   
                   'Define percent change
                   PercentChange = YearlyChange / OpeningPrice
                   
                   'Print percent change under percent change header as a percent
                   ws.Cells(row, 11).Value = Format(PercentChange, "percent")
                   
                   'Redefine opening price
                   OpeningPrice = ws.Cells(i + 1, 3).Value
                   
                   'Add one to the row number
                   row = row + 1
        
               End If
            
            'Otherwise (opening price is equal to 0), redefine opening price as price in following row.
            Else
            OpeningPrice = ws.Cells(i + 1, 3).Value
                
            End If
               
        'Move to next iteration
        Next i
        
        'Create table to include column headers "Ticker" and "Value."
        'Create row headers "Greatest % increase", "Greatest % Decrease" and "Greatest total volume".
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Create variables
        Dim Max As Variant
        Dim Min As Double
        Dim MaxTicker As String
        Dim MinTicker As String
        Dim MaxVolume As Variant
        Dim MaxVolumeTicker As String
        
        'Set initial max value to 0
        Max = 0
        
        'Set initial min value to 0
        Min = 0
        
        'Set initial max volume to 0
        MaxVolume = 0
    
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 11).End(xlUp).row
        
        'Loop through rows (looking for "Greatest % Increase" and "Greatest % Decrease")
        For i = 2 To LastRow
        
            'If cell's value is greater than the currently set Max
            If ws.Cells(i, 11).Value > Max Then
            
            'Then redefine Max with the current value
            Max = ws.Cells(i, 11).Value
            
            'Define the MaxTicker to be the ticker associated with the current Max
            MaxTicker = ws.Cells(i, 9).Value
            
            'If cell's value is less than the currently set Min
            ElseIf ws.Cells(i, 11).Value < Min Then
            
            'Then redefine Min with the current value
            Min = ws.Cells(i, 11).Value
            
            'Define the MinTicker to be the ticker associated with the current Min
            MinTicker = ws.Cells(i, 9).Value
        
            End If
        
        Next i
        
        'Determine the Last Row
        LastRow = ws.Cells(Rows.Count, 12).End(xlUp).row
        
        'Loop through rows (looking for "Greatest Total Volume")
        For i = 2 To LastRow
        
            'If cell's value is greater than the currently set max volume
            If ws.Cells(i, 12).Value > MaxVolume Then
            
            'Then redefine MaxVolume with the current value
            MaxVolume = ws.Cells(i, 12).Value
            
            'Define the MaxVolumeTicker to be the ticker associated with the current MaxVolume
            MaxVolumeTicker = ws.Cells(i, 9).Value
            
            End If
        
        Next i
        
        'Print the Max, Min, and MaxVolume along with the tickers associated in the appropriate cells.
        ws.Range("Q2").Value = Format(Max, "percent")
        ws.Range("Q3").Value = Format(Min, "percent")
        ws.Range("Q4").Value = MaxVolume
        ws.Range("P2").Value = MaxTicker
        ws.Range("P3").Value = MinTicker
        ws.Range("P4").Value = MaxVolumeTicker
        
        'Autofit all columns of active worksheet
        ws.Cells.Columns.AutoFit
        
    Next ws
    
End Sub
        













