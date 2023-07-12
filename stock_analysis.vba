Sub stock_analysis()
    
    ' Defines all variables used.
    Dim ticker_symbol As String
    Dim table_start As Double
    Dim lastRow As Long
    Dim i As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim total_table_row As Double
    Dim total_volume As Double
    Dim ws As Worksheet
       
    ' Goes through each worksheet in the workbook.
    For Each ws In ThisWorkbook.Worksheets
    
        'Setting the initial amount per variable.
        ticker_symbol = " "
        table_start = 2
        open_price = 0
        close_price = 0
        total_table_row = 2
        yearly_change = 0
        percent_change = 0
        total_volume = 0
        
        ' Placing the titles of the variables per column.
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        ' Setting each title to be bold.
        ws.Range("I1:L1").Font.Bold = True
        ws.Range("P1:Q1").Font.Bold = True
        ws.Range("O2:O4").Font.Bold = True
        
        ' Making sure the loop will go through the last row per sheet.
        lastRow = Range("A" & Rows.Count).End(xlUp).Row
    
        ' Inner loop for going through every row (starting at 2) to the last row.
        For i = 2 To lastRow
        
            ' Checks to see if we're still within the correct ticker symbol.
            ' If not - Writes the results to the summary table.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                ' Set ticker symbol.
                ticker_symbol = ws.Cells(i, 1).Value
                
                ' Defines the starting open/close prices.
                open_price = ws.Cells(table_start, 3).Value
                close_price = ws.Cells(i, 6).Value
                
                ' Calculates the yearly change.
                yearly_change = close_price - open_price
                
                
                ' If/Else statement regarding choosing appropriate color index per yearly change.
                If (yearly_change > 0) Then
                    ws.Range("J" & total_table_row).Interior.ColorIndex = 4
                
                ElseIf (yearly_change < 0) Then
                    ws.Range("J" & total_table_row).Interior.ColorIndex = 3
                    
                ElseIf (yearly_change = 0) Then
                    ws.Range("J" & total_table_row).Interior.ColorIndex = 0
                    
                End If
                
                
                ' Calculates the percent change.
                percent_change = (yearly_change / open_price) * 100
                
                ' Calculates the total volume to the table.
                total_volume = total_volume + ws.Cells(i, 7).Value
                
                
                ' Defines the range in which each value is placed.
                ws.Range("I" & total_table_row).Value = ticker_symbol
                ws.Range("J" & total_table_row).Value = yearly_change
                ws.Range("K" & total_table_row).Value = "%" & percent_change
                ws.Range("L" & total_table_row).Value = total_volume
                
                ' Resetting the values for the next loop iteration.
                total_table_row = total_table_row + 1
                table_start = i + 1
                yearly_change = 0
                percent_change = 0
                close_price = 0
                open_price = ws.Cells(table_start, 3).Value
                total_volume = 0
                
                
            Else
                total_volume = total_volume + ws.Cells(i, 7).Value
                
            End If
            
        Next i

        
        ' Taking the max/min percent change values and placing them in row Q.
        ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & lastRow)) * 100
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
        
        ' Making sure to not include the header row.
        increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)
        
        ' Puts the final ticker symbol per each total.
        ws.Range("P2") = ws.Cells(increase_number + 1, 9)
        ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
        ws.Range("P4") = ws.Cells(volume_number + 1, 9)
        
    Next ws
    
End Sub
