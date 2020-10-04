Sub alphabetical_testing()
    
    Dim WS_Count As Integer
    Dim J As Integer

    ' set ws count equal to number of worksheets
     WS_Count = ActiveWorkbook.worksheets.Count

    For J = 1 To WS_Count
         ' set stock ticker value
         Dim ticker As String
         ' set yearly change value
         Dim yearly_change As Double
         ' set yearly percent change value
         Dim yearly_percent As Double
         ' set total annual volume
         Dim total_volume As Double
         
         total_volume = 0
         
         Dim summary_table_row As Integer
         summary_table_row = 2
         
         Dim year_open_val As Double
         Dim year_end_val As Double
         Dim year_change As Double
         
         Dim max_increase As Double
             max_increase = 0
         Dim max_decrease As Double
             max_decrease = 0
         Dim max_vol As Double
             max_vol = 0
         
            
         numrows = Range("a2", Range("a2").End(xlDown)).Rows.Count
         
         
        For I = 2 To numrows
             
             If Cells(I, 2).Value = 20160101 Or Cells(I, 2).Value = 20150101 Or Cells(I, 2).Value = 20140101 Then
                 
                 year_open_val = Cells(I, 3).Value
             
             End If
              
             If Cells(I, 2).Value = 20161230 Or Cells(I, 2).Value = 20151230 Or Cells(I, 2).Value = 20141230 Then
             
                 year_end_val = Cells(I, 6).Value
             
             End If
                     
             total_volume = total_volume + Cells(I, 7).Value
        
             If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                 
                 year_change = year_end_val - year_open_val
                 
                 ticker = Cells(I, 1).Value
             
                 Range("I" & summary_table_row).Value = ticker
             
                 Range("L" & summary_table_row).Value = total_volume
             
                 Range("J" & summary_table_row).Value = year_change
             
                 Range("K" & summary_table_row).Value = (year_change / year_open_val)
                 
                 If Cells(summary_table_row, 11).Value > 0 Then
                     Cells(summary_table_row, 11).Interior.Color = vbGreen
                     Cells(summary_table_row, 11).NumberFormat = "0.00%"
                 Else
                     Cells(summary_table_row, 11).Interior.Color = vbRed
                     Cells(summary_table_row, 11).NumberFormat = "0.00%"
                 End If
                 
                 If total_volume > max_vol Then
                     max_vol = total_volume
                     Cells(4, 15).Value = ticker
                     Cells(4, 16).Value = max_vol
                     max_vol = total_volume
                 
                 End If
                 If total_volume < max_decrease Then
                     max_decrease = total_volume
                     Cells(3, 15).Value = max_decrease
                 End If
                             
                 summary_table_row = summary_table_row + 1
             
                 total_volume = 0
             
             End If
                 
                
        Next I

    Next J
End Sub

