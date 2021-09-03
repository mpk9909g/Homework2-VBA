Attribute VB_Name = "Module1"
Sub homework2()

For Each ws In Worksheets

    Dim worksheet_name As String
    worksheet_name = ws.Name
    MsgBox (worksheet_name)

' Define variables

    Dim ticker_column As Integer
    Dim price_open_column As Integer
    Dim price_close_column As Integer
    Dim volume_column As Integer
    Dim summary_row As LongLong
    Dim summary_ticker_col As Integer
    Dim summary_yearly_col As Integer
    Dim summary_percent_col As Integer
    Dim summary_volume_col As Integer
    Dim red_color As Integer
    Dim green_color As Integer


    Dim ticker As String
    Dim price_open As Double
    Dim price_close As Double
    Dim volume As LongLong

    Dim yearly_change As Double
    Dim percent_change As Double

    Dim last_row As LongLong
    Dim last_row_summary As LongLong


    Dim i As LongLong
    Dim j As LongLong

    'Set non-loop variables
    ticker_column = 1
    price_open_column = 3
    price_close_column = 6
    volume_column = 7

    summary_row = 2
    summary_ticker_col = 9
    summary_yearly_col = 10
    summary_percent_col = 11
    summary_volume_col = 12
    red_color = 3
    green_color = 4
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row


'initialize necessary loop variables
    price_open = ws.Cells(2, price_open_column).Value
    volume = 0



'Populate header for summary table

    ws.Cells(summary_row - 1, summary_ticker_col).Value = "Ticker"
    ws.Cells(summary_row - 1, summary_yearly_col).Value = "Yearly Change"
    ws.Cells(summary_row - 1, summary_percent_col).Value = "Percent Change"
    ws.Cells(summary_row - 1, summary_volume_col).Value = "Total Stock Volume"







'Loop for building the summary table

    For i = 2 To last_row

        ticker = ws.Cells(i, ticker_column).Value
        price_close = ws.Cells(i, price_close_column).Value
        yearly_change = price_close - price_open
        
        'include check to make sure opening price <>0
        If price_open = 0 Then
            percent_change = 0
        Else
            percent_change = price_close / price_open - 1
        End If
        
    
        volume = volume + ws.Cells(i, volume_column).Value


    
        'logic for identifying new ticker and populating summary table
        If ticker <> ws.Cells(i + 1, ticker_column).Value Then
        
           'set up summary table in the active sheet
        
            ws.Cells(summary_row, summary_ticker_col).Value = ticker
            ws.Cells(summary_row, summary_yearly_col).Value = yearly_change
            ws.Cells(summary_row, summary_percent_col).Value = percent_change
            ws.Cells(summary_row, summary_volume_col).Value = volume
      
            If yearly_change < 0 Then
        
                ws.Cells(summary_row, summary_yearly_col).Interior.ColorIndex = red_color
            
            Else
        
                ws.Cells(summary_row, summary_yearly_col).Interior.ColorIndex = green_color
            
            End If

        
        'reset necessary variables for the new ticker
            price_open = ws.Cells(i + 1, price_open_column).Value
            volume = 0
            summary_row = summary_row + 1
        
        End If

    
    Next i

'format the summary table

    ws.Columns("I:L").AutoFit
    ws.Columns("J:J").NumberFormat = "0.00"
    ws.Columns("K:K").NumberFormat = "0.00%"
    ws.Columns("l:l").NumberFormat = "0"


'now find the Greatest % increase, Greatest % decrease and Greatest total volume
    Dim ticker_pct_max As String
    Dim ticker_pct_min As String
    Dim ticker_vol_max As String
    Dim percent_change_max As Double
    Dim percent_change_min As Double
    Dim volume_max As LongLong

'Intialize / set the necessary variables pre-loop
    percent_change_max = 0
    percent_change_min = 0
    volume_max = 0

    last_row_summary = ws.Cells(Rows.Count, summary_ticker_col).End(xlUp).Row

'Loop for finding the max values
    For j = 2 To last_row_summary
        ticker = ws.Cells(j, summary_ticker_col).Value
        percent_change = ws.Cells(j, summary_percent_col).Value
        volume = ws.Cells(j, summary_volume_col).Value

        If percent_change > percent_change_max Then
            percent_change_max = percent_change
            ticker_pct_max = ticker
        ElseIf percent_change < percent_change_min Then
            percent_change_min = percent_change
            ticker_pct_min = ticker
        End If
   
        If volume > volume_max Then
            volume_max = volume
            ticker_vol_max = ticker
        End If

    Next j

'Create "Max" table
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O2").Value = ticker_pct_max
    ws.Range("O3").Value = ticker_pct_min
    ws.Range("O4").Value = ticker_vol_max
    ws.Range("P2").Value = percent_change_max
    ws.Range("P3").Value = percent_change_min
    ws.Range("P4").Value = volume_max

'format the max table

    ws.Columns("N:P").AutoFit

    ws.Range("P2:P3").NumberFormat = "0.00%"
    ws.Range("P4").NumberFormat = "0"

Next ws

End Sub



