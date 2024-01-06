Sub StockSummary()
    'Define Variables
    Dim ticker As String
    Dim year_open_price As Double
    Dim year_close_price As Double
    Dim vol As Double
    Dim max_vol As Double
    Dim max_vol_ticker As String
    Dim percent_change As Double
    Dim yearly_change As Double
    Dim max As Long
    Dim max_ticker As String
    Dim min As Long
    Dim min_ticker As String
    Dim ws As Worksheet
    Dim myrange As Range
    Dim condition1 As FormatCondition
    Dim condition2 As FormatCondition
    
    
    'Loop through all the worksheets in the workbook
    For Each ws In ThisWorkbook.worksheets
    
        'find last row for easier readability
        'used this stack overflow for https://stackoverflow.com/questions/11169445/find-last-used-cell-in-excel-vba
        Dim last_row As Long
        Dim last_row2 As Long
        last_row = Cells(Rows.Count, 1).End(xlUp).Row
        last_row2 = Cells(Rows.Count, 10).End(xlUp).Row
        
        
        'Add column headers and summary table
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'Initialize variables
        max = 0
        min = 0
        max_ticker = ""
        min_ticker = ""
        max_vol = 0
        max_vol_ticker = ""
            
        'nested loop to read each row in the current worksheet
        Dim ticker_row As Long
        ticker_row = 2
        For i = 2 To last_row
            
                'Check if the ticker is different from the previous row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                'Get the ticker symbol
                ticker = ws.Cells(i, 1).Value
                'Get the current years opening price
                year_open_price = ws.Cells(i, 3).Value
                'initialize/reset the volume variable
                vol = 0
            End If
            
            'Write the ticker symbol to the summary table
                ws.Cells(ticker_row, 9).Value = ticker
                
            'add the volume of the current row to the total volume
            vol = vol + ws.Cells(i, 7).Value
            
            'Check if the ticker is different from the next row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                'Get the current years closing price
                year_close_price = ws.Cells(i, 6).Value
                'Calculate the yearly change
                yearly_change = year_close_price - year_open_price
                'Write the yearly change to the summary table
                ws.Cells(ticker_row, 10).Value = "$" & yearly_change
                
                'calculate the percentage change. I had to add the conditional if the open price was 0 to avoid the instance that it would divide by zero.
                If year_open_price <> 0 Then
                    percent_change = (yearly_change / year_open_price) * 100
                Else
                    percent_change = 0
                End If
                'write the percentage change to the summary table
                ws.Cells(ticker_row, 11).Value = percent_change & "%"
                'write the total volume to the summary table
                ws.Cells(ticker_row, 12).Value = vol
                'Increment the summary row
                ticker_row = ticker_row + 1
                
                'Compare the percentage change with the maximum percentage
                If percent_change > max Then
                    'Update the maximum percentage and ticker
                    max = percent_change
                    max_ticker = ticker
                End If
                
                'Compare the percentage change with the minimum percentage
                If percent_change < min Then
                    'Update the minimum percentage and ticker
                    min = percent_change
                    min_ticker = ticker
                End If
                
                'Compare the volume with the maximum volume
                If vol > max_vol Then
                    'Update the maximum percentage and ticker
                    max_vol = vol
                    max_vol_ticker = ticker
                End If
            End If
        Next i
        
        'Write the greatest percentage increase/decrease and greatest total volume to the output cell
        ws.Cells(2, 16).Value = max_ticker
        ws.Cells(2, 17).Value = max & "%"
        
        ws.Cells(3, 16).Value = min_ticker
        ws.Cells(3, 17).Value = min & "%"
        
        ws.Cells(4, 16).Value = max_vol_ticker
        ws.Cells(4, 17).Value = max_vol
        
        'conditional formatting code I got from https://www.automateexcel.com/vba/conditional-formatting/
        'Set the range to apply conditional formatting
        Set myrange = ws.Range("J2:K" & last_row2)
        'Clear any existing conditional formatting
        myrange.FormatConditions.Delete
        'Add the first condition for values below 0
        Set condition1 = myrange.FormatConditions.Add(xlCellValue, xlLess, 0)
        'Set the format for the first condition
        With condition1
            .Interior.Color = vbRed
            .Font.Color = vbWhite
        End With
        'Add the second condition for values above 0
        Set condition2 = myrange.FormatConditions.Add(xlCellValue, xlGreater, 0)
        'Set the format for the second condition
        With condition2
            .Interior.Color = vbGreen
            .Font.Color = vbWhitee
        End With
        
    Next ws 'go to the next worksheet
End Sub