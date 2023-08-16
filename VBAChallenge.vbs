Sub MultipleYearStockData()

Dim ws As Worksheet
    Dim select_index As Double
    Dim ticker_row As Double
    Dim last_row As Double
    Dim year_opening As Double
    Dim year_closing As Double
    Dim total_stock_volume As Double

    
    For Each ws In Worksheets
        Worksheets(ws.Name).Activate
        select_index = 2
        ticker_row = 2
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        total_stock_volume = 0
        
        
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        
        For i = 2 To last_row
            tickers = ws.Cells(i, 1).Value
            tickers2 = ws.Cells(i - 1, 1).Value
            If tickers <> tickers2 Then
                ws.Cells(ticker_row, 9).Value = tickers
                ticker_row = ticker_row + 1
            End If
         Next i
    
        For i = 2 To last_row + 1
            tickers = ws.Cells(i, 1).Value
            tickers2 = ws.Cells(i - 1, 1).Value
            If tickers = tickers2 And i > 2 Then
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            ElseIf i > 2 Then
                ws.Cells(select_index, 12).Value = total_stock_volume
                select_index = select_index + 1
                total_stock_volume = 0
            Else
                total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
            End If
        Next i
            
        select_index = 2
        For i = 2 To last_row
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                year_closing = ws.Cells(i, 6).Value
            ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                year_opening = ws.Cells(i, 3).Value
            End If
            If year_opening > 0 And year_closing > 0 Then
                increase = year_closing - year_opening
                percent_increase = increase / year_opening
                ws.Cells(select_index, 10).Value = increase
                ws.Cells(select_index, 11).Value = FormatPercent(percent_increase)
                year_closing = 0
                year_opening = 0
                select_index = select_index + 1
            End If
        Next i
        
        
        max_per = WorksheetFunction.Max(ActiveSheet.Columns("k"))
        min_per = WorksheetFunction.Min(ActiveSheet.Columns("k"))
        max_vol = WorksheetFunction.Max(ActiveSheet.Columns("l"))
        
        ws.Range("Q2").Value = FormatPercent(max_per)
        ws.Range("Q3").Value = FormatPercent(min_per)
        ws.Range("Q4").Value = max_vol
        
        
        For i = 2 To last_row
            If max_per = ws.Cells(i, 11).Value Then
                ws.Range("P2").Value = ws.Cells(i, 9).Value
            ElseIf min_per = ws.Cells(i, 11).Value Then
                ws.Range("P3").Value = ws.Cells(i, 9).Value
            ElseIf max_vol = ws.Cells(i, 12).Value Then
                ws.Range("P4").Value = ws.Cells(i, 9).Value
            End If
        Next i
        
        
        For i = 2 To last_row
            If ws.Cells(i, 10).Value > 0 Then
                ws.Cells(i, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 10).Interior.ColorIndex = 3
            End If
        Next i
        
        For i = 2 To last_row
            If ws.Cells(i, 11).Value > 0 Then
                ws.Cells(i, 11).Interior.ColorIndex = 4
            Else
                ws.Cells(i, 11).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
                
End Sub