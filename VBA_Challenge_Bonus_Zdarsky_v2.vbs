Attribute VB_Name = "Module1"
Sub RunCode()
    'Inlcudes Code Bonus Section
    
    Dim WS_Count As Integer
    Dim I As Integer
    Dim J As Long
    Dim Tickers_in_ws As Integer
    ' Set a variable for specifying the column/row of interest
    Dim column As Integer
    Dim lastrow As Double
    Dim lastcolumn As Double
    Dim Total_volume As Double
    Dim Initial_bonus_counter As Integer            'For Bonus Section
    Dim Greatest_increase As Double                 'For Bonus Section
    Dim Greatest_decrease As Double                 'For Bonus Section
    Dim Greatest_total_volume As Double             'For Bonus Section
    
    'Debug.Print "Start of Program", Format(Now, "mm/dd/yyyy HH:mm:ss")
    Tickers_in_ws = 2           'initial row where store Ticker data in a worksheet
    Initial_bonus_counter = 0   'intialize counter for Bounus section
        
    ' Set WS_Count equal to the number of worksheets in the active
    ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
    For I = 1 To WS_Count
    'For I = 1 To 1
       'loop for worksheets

       'Debug.Print "ActiveWorkbook.Worksheets(i).Name is ", ActiveWorkbook.Worksheets(I).Name
        
      column = 1
      Total_volume = 0
      Opening_price = ActiveWorkbook.Worksheets(I).Cells(2, 3).Value
      Closing_price = 0
      lastrow = ActiveWorkbook.Worksheets(I).Cells(Rows.Count, 1).End(xlUp).Row
      'Debug.Print "lastrow=", lastrow
      'Set-up Column Headers
      ActiveWorkbook.Worksheets(I).Cells(1, 9).Value = "Ticker"
      ActiveWorkbook.Worksheets(I).Cells(1, 10).Value = "Yearly Change"
      ActiveWorkbook.Worksheets(I).Cells(1, 11).Value = "Percent Change"
      ActiveWorkbook.Worksheets(I).Cells(1, 12).Value = "Total Stock Volume"
      ActiveWorkbook.Worksheets(I).Cells(1, 16).Value = "Ticker"                'For Bonus Section
      ActiveWorkbook.Worksheets(I).Cells(1, 17).Value = "Value"                 'For Bonus Section
    
      'Loop through rows in the column
      'For J = 2 To 1049
      For J = 2 To lastrow
        Total_volume = Total_volume + ActiveWorkbook.Worksheets(I).Cells(J, 7).Value
    
        ' Searches for when the value of the next cell is different than that of the current cell
        If ActiveWorkbook.Worksheets(I).Cells(J + 1, column).Value <> ActiveWorkbook.Worksheets(I).Cells(J, column).Value Then
    
          Closing_price = ActiveWorkbook.Worksheets(I).Cells(J, 6).Value         'Set closing price
          Yearly_change = Closing_price - Opening_price
          If Yearly_change = 0 Then  'Handles when no change (e.g. Test data had zeros for ticker PLNT)
            Percent_change = 0
          Else
            If Opening_price = 0 Then
                Percent_change = 0
            Else
                Percent_change = Yearly_change / Opening_price
            End If
          End If
          
          'Debug.Print ActiveWorkbook.Worksheets(I).Cells(J, 1).Value, Yearly_change, Percent_change, Total_volume
          
          ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 9).Value = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value     'Store Ticker
          ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 10).Value = Yearly_change       'Store Yearly Change
          If Yearly_change < 0 Then
            ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 10).Interior.ColorIndex = 3
          Else
            ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 10).Interior.ColorIndex = 4
          End If
          ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 11).Value = Percent_change      'Store Percent Change
          ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 12).Value = Total_volume        'Store Total Stock Volume
          If Initial_bonus_counter = 0 Then         'Initialize values for Bonus Section
            'Debug.Print "In Initial Bonus Counter Loop"
            Greatest_increase = ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 11).Value         'Percent_change
            Greatest_increase_ticker = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value
            Greatest_decrease = ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 11).Value         'Percent_change
            Greatest_decrease_ticker = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value
            Greatest_total_volume = ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 12).Value
            Greatest_total_volume_ticker = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value
            Initial_bonus_counter = Initial_bonus_counter + 1
            'Debug.Print "Greatest_increase, Greatest_decrease, Greatest_total_volume", Greatest_increase, Greatest_decrease, Greatest_total_volume
          Else                                      'Determine Greatest values for Bonus Section
            If Greatest_increase < ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 11).Value Then
                Greatest_increase = ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 11).Value         'Percent_change
                Greatest_increase_ticker = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value
            End If
            If Greatest_decrease > ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 11).Value Then
                Greatest_decrease = ActiveWorkbook.Worksheets(I).Cells(Tickers_in_ws, 11).Value         'Percent_change
                Greatest_decrease_ticker = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value
            End If
            If Greatest_total_volume < Total_volume Then
                Greatest_total_volume = Total_volume
                Greatest_total_volume_ticker = ActiveWorkbook.Worksheets(I).Cells(J, 1).Value
            End If
            Initial_bonus_counter = Initial_bonus_counter + 1
            'Debug.Print "Greatest_increase, Greatest_decrease, Greatest_total_volume", Greatest_increase, Greatest_decrease, Greatest_total_volume
          End If
        
    
          Total_volume = 0                                                                  'Reset Total_volume for next stock ticker
          Opening_price = ActiveWorkbook.Worksheets(I).Cells(J + 1, 3).Value                'Set starting price for next ticker
        
        Tickers_in_ws = Tickers_in_ws + 1
        End If
        
      Next J
    
    ActiveWorkbook.Worksheets(I).Range("J2:J" & Tickers_in_ws).NumberFormat = "0.00"        'Set column J format to Number
    ActiveWorkbook.Worksheets(I).Range("K2:K" & Tickers_in_ws).NumberFormat = "0.00%"       'Set column K format to %
    ActiveWorkbook.Worksheets(I).Range("L:L").Columns.AutoFit                               'Set column L width to Autofit
    
    ActiveWorkbook.Worksheets(I).Cells(2, 15).Value = "Greatest % Increase"                 'Greatest Increase Label For Bonus Section
    ActiveWorkbook.Worksheets(I).Cells(3, 15).Value = "Greatest % Decrease"                 'Greatest Decrease Label For Bonus Section
    ActiveWorkbook.Worksheets(I).Cells(4, 15).Value = "Greatest Total Volume"               'Greatest Total Volume Label For Bonus Section
    
    ActiveWorkbook.Worksheets(I).Cells(2, 16).Value = Greatest_increase_ticker              'Greatest Increase Ticker For Bonus Section
    ActiveWorkbook.Worksheets(I).Cells(3, 16).Value = Greatest_decrease_ticker              'Greatest Decrease Ticker For Bonus Section
    ActiveWorkbook.Worksheets(I).Cells(4, 16).Value = Greatest_total_volume_ticker          'Greatest Total Volume Ticker For Bonus Section
    
    ActiveWorkbook.Worksheets(I).Cells(2, 17).Value = Greatest_increase                     'Greatest Increase For Bonus Section
    ActiveWorkbook.Worksheets(I).Cells(3, 17).Value = Greatest_decrease                     'Greatest Decrease For Bonus Section
    ActiveWorkbook.Worksheets(I).Cells(4, 17).Value = Greatest_total_volume                 'Greatest Total Volume For Bonus Section
    
    ActiveWorkbook.Worksheets(I).Range("O:O").Columns.AutoFit                               'Set column O width to Autofit For Bonus Section
    ActiveWorkbook.Worksheets(I).Range("Q2:Q3").NumberFormat = "0.00%"                      'Set to % For Bonus Section
    'ActiveWorkbook.Worksheets(I).Range("Q4:Q4").NumberFormat = "0"                          'Set to Number For Bonus Section

    Tickers_in_ws = 2                                                                       'Reset Tickers_in_ws to start in second row
    Initial_bonus_counter = 0                                                               'Intialize counter for Bounus section
    
    Next I
    'Debug.Print "Program Ended", Format(Now, "mm/dd/yyyy HH:mm:ss")
End Sub


