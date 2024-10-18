Attribute VB_Name = "Module1"
Sub multiple_Quarter_stock_data()

   'set variables
Dim ticker As String
Dim number_tickers As Integer
Dim lastRowState As Long
Dim opening_price As Double
Dim closing_price As Double
Dim quarterly_change As Double
Dim percent_change As Double
Dim total_stock_volume As Double
Dim greatest_percent_increase As Double
Dim greatest_percent_increase_ticker As String
Dim greatest_percent_decrease As Double
Dim greatest_percent_decrease_ticker As String
Dim greatest_stock_volume As Double
Dim greatest_stock_volume_ticker As String

' looping and activating worksheets
For Each ws In Worksheets

    ws.Activate

    ' Find the last row
    lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ' add headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Quarterly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ' set variables to 0
    number_tickers = 0
    ticker = ""
    quarterly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
    ' loop through tickers
    For i = 2 To lastRowState
        
        ticker = Cells(i, 1).Value
        
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
        ' Adding  the total stock volume
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        ' if we get different ticker
        If Cells(i + 1, 1).Value <> ticker Then
            number_tickers = number_tickers + 1
            Cells(number_tickers + 1, 9) = ticker
            
            ' closing price 
            closing_price = Cells(i, 6)
            
            ' Change value
            quarterly_change = closing_price - opening_price
            
            Cells(number_tickers + 1, 10).Value = quarterly_change
            
            ' greater than 0 = green.
            If quarterly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ' less than 0 = red.
            ElseIf quarterly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            'is 0 = yellow.
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
            
            
            ' Calculate percent change
            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (quarterly_change / opening_price)
            End If
            
            
            ' Format the percent_change value
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
            
            ' setting the opening price back to 0
            opening_price = 0
            
            ' Add total stock value
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
            ' Setting total stock volume back to 0
            total_stock_volume = 0
        End If
        
    Next i
    
    '2nd. greatest percent increase, greatest percent decrease, and greatest total volume.
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    ' Geting  last row
    lastRowState = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ' get variables
    greatest_percent_increase = Cells(2, 11).Value
    greatest_percent_increase_ticker = Cells(2, 9).Value
    greatest_percent_decrease = Cells(2, 11).Value
    greatest_percent_decrease_ticker = Cells(2, 9).Value
    greatest_stock_volume = Cells(2, 12).Value
    greatest_stock_volume_ticker = Cells(2, 9).Value
    
    
    ' loop through tickers.
    For i = 2 To lastRowState
    
        '  greatest percent increase.
        If Cells(i, 11).Value > greatest_percent_increase Then
            greatest_percent_increase = Cells(i, 11).Value
            greatest_percent_increase_ticker = Cells(i, 9).Value
        End If
        
        ' greatest percent decrease.
        If Cells(i, 11).Value < greatest_percent_decrease Then
            greatest_percent_decrease = Cells(i, 11).Value
            greatest_percent_decrease_ticker = Cells(i, 9).Value
        End If
        
        ' greatest stock volume.
        If Cells(i, 12).Value > greatest_stock_volume Then
            greatest_stock_volume = Cells(i, 12).Value
            greatest_stock_volume_ticker = Cells(i, 9).Value
        End If
        
    Next i
    
    ' Adding the values for greatest percent increase, decrease, and stock volume
    Range("P2").Value = Format(greatest_percent_increase_ticker, "Percent")
    Range("Q2").Value = Format(greatest_percent_increase, "Percent")
    Range("P3").Value = Format(greatest_percent_decrease_ticker, "Percent")
    Range("Q3").Value = Format(greatest_percent_decrease, "Percent")
    Range("P4").Value = greatest_stock_volume_ticker
    Range("Q4").Value = greatest_stock_volume
    
Next ws

        
End Sub
