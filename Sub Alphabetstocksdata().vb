Sub Alphabetstocksdata()
 Dim ws As Worksheet
    For Each ws In Worksheets
        'define the variables in the data
    
        Dim i As Long
        Dim j As Long
        Dim ticker As String
        Dim number_tickers As Integer
        Dim last_row_stat As Long
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim total_stock_volume As Double
        Dim percent_change As Double
        Dim greatest_percent_increase As Double
        Dim greatest_percent_decrease As Double
        Dim greatest_total_volume As Double
        
'name the columns that will hold the data values we need for all worksheets
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
 
 'set up integers for the loop
        ticker = 2
        i = 2
        j = 2
        LastRowl = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To LastRowl
    
       'condition for when the ticker symbol is the same if not then another function
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'Write ticker in column I
                ws.Cells(ticker, 9).Value = ws.Cells(i, 1).Value
                
                'Calculate Yearly Change in column J (close-open)
                ws.Cells(ticker, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                
                    'Conditional formating to change color
                    If ws.Cells(ticker, 10).Value < 0 Then
                
                    'Set cell background color to red
                    ws.Cells(ticker, 10).Interior.ColorIndex = 3
                
                    Else
                
                    'Set cell background color to green
                    ws.Cells(ticker, 10).Interior.ColorIndex = 4
                
                    End If
                    
                    'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                    percent_change = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    
                    'Percent formating
                    ws.Cells(ticker, 11).Value = Format(percent_change, "Percent")
                    
                    Else
                    
                    ws.Cells(ticker, 11).Value = Format(0, "Percent")
                    
                    End If
                    
                'Calculate and write total volume in column L (#12)
                ws.Cells(ticker, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                
                'Increase TickCount by 1
                ticker = ticker + 1
                
                'Set new start row of the ticker block
                j = i + 1
                
                End If
            
            Next i
            
        'Find last non-blank cell in column I
        LastRowI = ws.Cells(Rows.Count, 9).End(xlUp).Row
        
        'Summary table
        greatest_total_volume = ws.Cells(2, 12).Value
        greatest_percent_increase = ws.Cells(2, 11).Value
        greatest_percent_decrease = ws.Cells(2, 11).Value
        
            'Loop for summary
            For i = 2 To LastRowI
            
                'For greatest total volume--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 12).Value > greatest_total_volume Then
                greatest_total_volume = ws.Cells(i, 12).Value
                ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatest_total_volume = greatest_total_volume
                
                End If
                
                'For greatest increase--check if next value is larger--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value > greatest_percent_increase Then
                greatest_percent_increase = ws.Cells(i, 11).Value
                ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatest_percent_increase = greatest_percent_increase
                
                End If
                
                'For greatest decrease--check if next value is smaller--if yes take over a new value and populate ws.Cells
                If ws.Cells(i, 11).Value < greatest_percent_decrease Then
                greatest_percent_decrease = ws.Cells(i, 11).Value
                ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                
                Else
                
                greatest_percent_decrease = greatest_percent_decrease
                
                End If
                
            'Write summary results in ws.Cells
            ws.Cells(2, 17).Value = Format(greatest_percent_increase, "Percent")
            ws.Cells(3, 17).Value = Format(greatest_percent_decrease, "Percent")
            ws.Cells(4, 17).Value = Format(greatest_total_volume, "Scientific")
            
            Next i
     
        
            
    Next ws
        
End Sub

