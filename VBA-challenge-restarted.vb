
Sub Summarize_StocksYearly()

     Dim WS As Worksheet

     For Each WS In ThisWorkbook.Sheets
     
        'Set to activate each worksheet
        WS.Activate
        
        MsgBox "Processing worksheet: " & WS.Name
        
        Dim LastRow As Long

        ' Define Ticker and it's column variables for calculations.
        Dim ticker_name As String
        Dim total_tradeday_volume As Double
        Dim year_open_price As Double
        Dim year_close_price As Double
        Dim greatest_increase_during_yr As Double
        Dim greatest_decrease_during_yr As Double
        Dim highest_total_volume As Double
        Dim year_change_calc As Double
        Dim percent_change_during_year As Double
     
        ' Confirmed the data headers of the input data + establish reference for output needs to summary table.
        WS.Cells(1, 9).Value = "Ticker"
        WS.Cells(1, 10).Value = "Yearly Change"
        WS.Cells(1, 11).Value = "Percent Change"
        WS.Cells(1, 12).Value = "Total Stock Volume"
     
        ' Set the headers of the Greatest Changes Summary Table.
        WS.Cells(2, 15).Value = "Greatest % Increase"
        WS.Cells(3, 15).Value = "Greatest % Decrease"
        WS.Cells(4, 15).Value = "Greatest Total Volume"
        WS.Cells(1, 16).Value = "Ticker"
        WS.Cells(1, 17).Value = "Value"
     
        'Establish target points for looper to start with as reference.
        year_first_record = 2
        currRow_start = 2
            
        ' Determine the Last Row
        lastRow_end = WS.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Determine the years first record
        year_open_price = WS.Cells(year_first_record, 3).Value
               
        'Loop through all ticker's type and stock volumes
        'i = 2
        For i = 2 To lastRow_end

           ' After reaching the last row of the current ticker symbol data
            If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
                ticker_name = WS.Cells(i, 1).Value
                WS.Cells(currRow_start, 9).Value = ticker_name
                
                ' Set year's close price
                year_close_price = WS.Cells(i, 6).Value
                
                ' Calculating the yearly change between the year close and year open
                year_change_calc = (year_close_price - year_open_price)
                WS.Cells(currRow_start, 10).Value = year_change_calc
                
                ' Calcuating the percent change between the year close and year open
                percent_change_during_year = (year_change_calc / year_open_price)
                WS.Cells(currRow_start, 11).Value = percent_change_during_year
                WS.Cells(currRow_start, 11).Style = "Percent"
                
                ' Add volume of row to total volume
                total_tradeday_volume = total_tradeday_volume + WS.Cells(i, 7).Value
                WS.Cells(currRow_start, 12).Value = total_tradeday_volume
                                            
                ' Increase the row of the summary table for next ticker information
                currRow_start = currRow_start + 1
            
                ' Reset total volume for the next ticker data
                total_tradeday_volume = 0
                
                year_change_calc = 0
                
                year_open_price = WS.Cells(i + 1, 3).Value
            
            ' While looping through data of the current ticker symbol
            Else
                ' Add volume of row to total volume
                total_tradeday_volume = total_tradeday_volume + WS.Cells(i, 7).Value
            
            End If
       Next i
       

       greatest_increase_during_yr = 0
       greatest_decrease_during_yr = Cells(2, 11)
       highest_total_volume = 0
       lastRow_end2 = WS.Cells(Rows.Count, 9).End(xlUp).Row
       
       ' Loop through the Yearly change data and color the cells based on + / -
       For j = 2 To lastRow_end2
       
        If WS.Cells(j, 10) >= 0 Then
            WS.Cells(j, 10).Interior.ColorIndex = 4
        
        Else
            WS.Cells(j, 10).Interior.ColorIndex = 3
        
        End If
           
        If WS.Cells(j, 11) > greatest_increase_during_yr Then
            greatest_increase_during_yr = WS.Cells(j, 11).Value
                  
            WS.Cells(2, 17).Value = greatest_increase_during_yr
            WS.Cells(2, 17).Style = "Percent"
            WS.Cells(2, 16).Value = Cells(j, 9).Value
        
        ElseIf WS.Cells(j, 11) < greatest_decrease_during_yr Then
            greatest_decrease_during_yr = WS.Cells(j, 11).Value
        
            WS.Cells(3, 17).Value = greatest_decrease_during_yr
            WS.Cells(3, 17).Style = "Percent"
            WS.Cells(3, 16).Value = WS.Cells(j, 9).Value
        
        End If
        
        If WS.Cells(j, 12).Value > highest_total_volume Then
            highest_total_volume = WS.Cells(j, 12).Value
        
            WS.Cells(4, 17).Value = highest_total_volume
            WS.Cells(4, 16).Value = Cells(j, 9).Value
        
        End If
       
       Next j
       
    Next WS
    
    MsgBox ("All Done!")

End Sub