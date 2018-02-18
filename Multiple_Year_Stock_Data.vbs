Sub Multiple_Year_Stock_Data()
    
    'Loop through all sheets
    For Each ws In Worksheets
    
        Dim summary_table_row As Integer
        summary_table_row = 2
        
        Dim ticker_opened_price As Double
        ticker_opened_price = ws.Cells(2, 3).Value
        
        Dim totalTickerVolume As Double
        totalTickerVolume = 0
        
        Dim ticker_closed_price As Double
        Dim ticker_yearly_change As Double
        Dim ticker_percent_change As Double
        
        Dim greatest_percent_increase As Double
        greatest_percent_increase = 0
        
        Dim greatest_repcent_decrease As Double
        greatest_repcent_decrease = 0
        
        Dim greatest_total_volume As Double
        greatest_total_volume = 0
        
        Dim greatest_increase_ticker As String
        Dim greatest_decrease_ticker As String
        Dim greatest_total_ticker As String
        
       'Name Summary column
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        'Create variable to hold ticker name, last row
        Dim ticker_name As String
        
        Dim lastRow As Long
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
                
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    ticker_name = ws.Cells(i, 1).Value
                    
                    
                    totalTickerVolume = totalTickerVolume + ws.Cells(i, 7).Value
                    
                    ws.Range("I" & summary_table_row).Value = ticker_name
                    
                    ws.Range("L" & summary_table_row).Value = totalTickerVolume
                    
                    
                       If totalTickerVolume > greatest_total_volume Then
                           greatest_total_volume = totalTickerVolume
                           greatest_total_ticker = ws.Cells(i, 1).Value
                       End If
                    
                   
                    ticker_closed_price = ws.Cells(i, 6).Value
                    
                    ticker_yearly_change = ticker_closed_price - ticker_opened_price
                    ws.Range("J" & summary_table_row).Value = ticker_yearly_change
                    
                        If ws.Range("J" & summary_table_row).Value > 0 Then
                            ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                        Else
                            ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                        End If
                        
                         If ticker_opened_price <> 0 Then
                             ticker_percent_change = ticker_yearly_change / ticker_opened_price
                        Else
                             ticker_percent_change = ticker_yearly_change / ticker_yearly_change
                         End If
                    
                        If ticker_percent_change > greatest_percent_increase Then
                            greatest_percent_increase = ticker_percent_change
                            greatest_increase_ticker = ws.Cells(i, 1).Value
                        
                        ElseIf ticker_percent_change < greatest_percent_decrease Then
                            greatest_percent_decrease = ticker_percent_change
                            greatest_decrease_ticker = ws.Cells(i, 1).Value
                        End If
                        
                        
                    ws.Range("K" & summary_table_row).Value = ticker_percent_change
            
                    ws.Range("K" & summary_table_row).NumberFormat = "0.00%"
                    
                    ticker_opened_price = ws.Cells(i + 1, 3).Value
                    
                    summary_table_row = summary_table_row + 1
                    totalTickerVolume = 0
                    
                Else
                    totalTickerVolume = totalTickerVolume + ws.Cells(i, 7).Value
                    
                End If
            Next i
            
            ws.Range("P2").Value = greatest_increase_ticker
            ws.Range("P3").Value = greatest_decrease_ticker
            ws.Range("Q2").Value = greatest_percent_increase
            ws.Range("Q2").NumberFormat = "0.00%"
            ws.Range("Q3").Value = greatest_percent_decrease
            ws.Range("Q3").NumberFormat = "0.00%"
            ws.Range("P4").Value = greatest_total_ticker
            ws.Range("Q4").Value = greatest_total_volume
                
      Next ws
  
End Sub
