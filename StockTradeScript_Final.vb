Sub Stocktrade()


Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate

    
            
        
            
LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

    Dim ticker As String
    Dim ticker_row As Long
    Dim stock_row As Long
    Dim stock_volume As String
    Dim diff As Variant
    Dim open_value As Variant
    Dim close_value As Variant
    Dim diff_row As Long
    Dim percent_change As Double
    Dim percent_row As Long
    
   
    
    
    
 
 

    
    
  
 

    
    
        ticker_row = 2
        stock_row = 2
        stock_volume = 0
        volume_row = 2
        diff_row = 2
        open_row = 2
        percent_row = 2
        open_counter = 0
        
       
        
        
       
        
        
        
 
 Cells(1, 9).Value = "Ticker"
 
 Cells(1, 10).Value = "Yearly Change"
 
 Cells(1, 11).Value = "Percent Change"
 
 Cells(1, 12).Value = "Total Stock Volume"
 
 
 
Worksheets("2014").Columns("A:P").AutoFit
Worksheets("2015").Columns("A:P").AutoFit
Worksheets("2016").Columns("A:P").AutoFit

 
        
        

        
    
On Error Resume Next


'Begin loop

For i = 2 To LastRow

    If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    
        stock_volume = stock_volume + WS.Cells(i + 1, 7)
    
            open_counter = open_counter + 1
    
            open_value = WS.Cells((i + 1) - open_counter, 3).Value
    
            close_value = WS.Cells(i, 6).Value
    
            diff = close_value - open_value
                       
            WS.Cells(diff_row, 10).Value = diff
            
            percent_change = Cells(diff_row, 10).Value / open_value
            
            WS.Cells(diff_row, 11).Value = percent_change
            
            WS.Cells(diff_row, 11).NumberFormat = "0.00%"
    
    
    
    
    ticker = Cells(i - 1, 1).Value
    
    WS.Cells(ticker_row, 9).Value = ticker
    
    WS.Cells(volume_row, 12).Value = stock_volume
    
    volume_row = volume_row + 1
    
    ticker_row = ticker_row + 1
    
    percent_row = percent_row + 1
    
    diff_row = diff_row + 1

    stock_volume = 0
    
    open_counter = 0
    
    
    

    
        
Else

    stock_volume = stock_volume + WS.Cells(i + 1, 7).Value
    
    open_counter = open_counter + 1
    
    
    
    
End If


Next i

'''' Color the cells

For i = 2 To LastRow

    If Cells(i, 10).Value > 0 Then

        Cells(i, 10).Interior.ColorIndex = 4

    ElseIf Cells(i, 10).Value < 0 Then

        Cells(i, 10).Interior.ColorIndex = 3

    ElseIf Cells(i, 10).Value = "" Then

        Cells(i, 10).Interior.ColorIndex = 0
    

End If

Next i


Next WS


''''



End Sub























