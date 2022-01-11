Attribute VB_Name = "Module1"
Sub RunCode()
    
    Dim WS_Count As Integer
    Dim I As Integer
    Dim J As Long
    Dim Tickers_in_ws As Integer
    ' Set a variable for specifying the column/row of interest
    Dim column As Integer
    Dim lastrow As Double
    Dim lastcolumn As Double
    Dim Total_volume As Double
    
    'Debug.Print "Start of Program", Format(Now, "mm/dd/yyyy HH:mm:ss")
    Tickers_in_ws = 2           'initial row where store Ticker data in a worksheet
    
        
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
          
    
          Total_volume = 0                                                                  'Reset Total_volume for next stock ticker
          Opening_price = ActiveWorkbook.Worksheets(I).Cells(J + 1, 3).Value                'Set starting price for next ticker
        
        Tickers_in_ws = Tickers_in_ws + 1
        End If
        
      Next J
    
    ActiveWorkbook.Worksheets(I).Range("J2:J" & Tickers_in_ws).NumberFormat = "0.00"        'Set column J format to Number
    ActiveWorkbook.Worksheets(I).Range("K2:K" & Tickers_in_ws).NumberFormat = "0.00%"       'Set column K format to %
    ActiveWorkbook.Worksheets(I).Range("L:L").Columns.AutoFit                               'Set column L width to Autofit
    

    Tickers_in_ws = 2                                                                       'Reset Tickers_in_ws to start in second row

    
    Next I
    'Debug.Print "Program Ended", Format(Now, "mm/dd/yyyy HH:mm:ss")
End Sub
