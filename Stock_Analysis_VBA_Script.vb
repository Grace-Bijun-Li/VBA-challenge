Sub Stock_Analysis():
'loop through and grab values for tickers

For Each ws In Worksheets


    Dim Total_stock_volume As Variant
    Summary_table_row = 2
    Open_row = 2
    
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
 
    'To calculate Max or Min percent change
    Dim x, y As Double
    x = -1
    y = 1
    
    Dim max_ticker, min_ticker As String
    
    
    'To calculate Max total volume
    Dim z As Variant
    z = -1
    
    Dim max_volume_ticker As String
    
    
    
    'Loop through and grab values for ticker
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    For Row = 2 To LastRow
    
     
        'Input the current values in Summary Table
        Total_stock_volume = Total_stock_volume + ws.Cells(Row, 7).Value
        

            If ws.Cells(Row, 1).Value <> ws.Cells(Row + 1, 1).Value Then
            
                'Grab new Ticker to Summary Table
                Ticker = ws.Cells(Row, 1).Value
                ws.Cells(Summary_table_row, 9).Value = Ticker
    
                'Calculate Yearly Change
                Open_value = ws.Cells(Open_row, 3).Value
                Close_value = ws.Cells(Row, 6).Value
                Yearly_change = Close_value - Open_value
                ws.Cells(Summary_table_row, 10).Value = Yearly_change
    
                'Calculate Percent Change
                    If (Open_value = 0) Then
                    ws.Cells(Summary_table_row, 11).Value = 0
                    Else
                    ws.Cells(Summary_table_row, 11).Value = Yearly_change / Open_value
                    End If
                        
                'Calculate Max or Min percent change
                new_x = WorksheetFunction.Max((ws.Cells(Summary_table_row, 11).Value), x)
                    If new_x > x Then
                        x = new_x
                        max_ticker = Ticker
                    End If
                    
                new_y = WorksheetFunction.Min((ws.Cells(Summary_table_row, 11).Value), y)
                    If new_y < y Then
                        y = new_y
                        min_ticker = Ticker
                    End If
                        
                'Format Percent Change
                ws.Cells(Summary_table_row, 11).NumberFormat = "0.00%"
                        
                'Calculate Total_volume
                ws.Cells(Summary_table_row, 12).Value = Total_stock_volume
                
                'Calculate Max total volume
                new_z = WorksheetFunction.Max((ws.Cells(Summary_table_row, 12).Value), z)
                    If new_z > z Then
                        z = new_z
                        max_volume_ticker = Ticker
                    End If
                               
                'Format the total stock value
                ws.Cells(Summary_table_row, 12).NumberFormat = "#,##0"
                    
                'Color the Summary Table
                    If Yearly_change <= 0 Then
                
                        'Color red
                        ws.Cells(Summary_table_row, 10).Interior.Color = RGB(255, 0, 0)
                    Else
                        'Color green
                        ws.Cells(Summary_table_row, 10).Interior.Color = RGB(0, 255, 0)
                    End If
    

                'Reset Summary Table for new ticker
                Total_stock_volume = 0
                Summary_table_row = Summary_table_row + 1
                Open_row = Row + 1
    
            End If

    Next Row


    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    'Put Max percent change value in the second summary table
    ws.Range("P2").Value = max_ticker
    ws.Range("Q2").Value = x
    
    'Put Min percent change value in the second summary table
    ws.Range("P3").Value = min_ticker
    ws.Range("Q3").Value = y
    
    'Format Percent Change in second summary table
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    
    'Put Max total volume in the second summary table
    ws.Range("P4").Value = max_volume_ticker
    ws.Range("Q4").Value = z
    
    
Next ws

End Sub
