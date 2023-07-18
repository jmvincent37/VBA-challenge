Sub vba_challenge()

Dim ticker_name As String
Dim Volume As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Summary_Table As Integer
Dim Rng As Range
Dim Ticker_Next As String



For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    
    
    ws.Cells(1, "I").Value = "Ticker"
    ws.Cells(1, "J").Value = "Yearly Change"
    ws.Cells(1, "K").Value = "Percent Change"
    ws.Cells(1, "L").Value = "Total Stock Volume"
    ws.Cells(2, "O").Value = "Greatest % Increase"
    ws.Cells(3, "O").Value = "Greatest % Decrease"
    ws.Cells(4, "O").Value = "Greatest Total Volume"
    ws.Cells(2, "P").Value = percentMaxTicker
    ws.Cells(3, "P").Value = percentMinTicker
    ws.Cells(4, "P").Value = volumeMaxTicker
    ws.Cells(2, "Q").Value = Format(PercentMax, "#.##%")
    ws.Cells(3, "Q").Value = Format(PercentMin, "#.##%")
    ws.Cells(4, "Q").Value = volumeMax
    ws.Columns("K").NumberFormat = "0.00%"
        
    
    RowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    
   
    Volume = 0
    greatest_increase = 0
    gi_ticker = ""
    greatest_decrease = 0
    gd_ticker = ""
    greatest_volume = 0
    gv_ticker = ""
    
    Summary_Table_Row = 2
    Open_Price_Pointer = 2
    
        For i = 2 To RowCount
        
                                      
             
             
                Volume = Volume + ws.Cells(i, "G").Value
                If ws.Cells(i, "A").Value <> ws.Cells(i - 1, "A").Value Then
                    Open_Price = ws.Cells(i, 3).Value
                    
            ElseIf ws.Cells(i, "A").Value <> ws.Cells(i + 1, "A").Value Then
                
            
                 
                 
    
                Close_Price = ws.Cells(i, "F").Value
                Yearly_Change = (Close_Price - Open_Price)
                
                If (Open_Price = 0) Then
                    Percent_Change = Yearly_Change
                Else
                    Percent_Change = Yearly_Change / Open_Price
                End If
                
                If Percent_Change > greatest_increase Then
                
                    greatest_increase = Percent_Change
                    gi_ticker = ws.Cells(i, "A").Value
                End If
                
                 If Percent_Change < greatest_decrease Then
                    greatest_decrease = Percent_Change
                    gd_ticker = ws.Cells(i, "A").Value
                End If
                
                If Volume > greatest_volume Then
                    greatest_volume = Volume
                    gv_ticker = ws.Cells(i, "A").Value
                End If
                       
                ws.Cells(Summary_Table_Row, "I").Value = ws.Cells(i, "A").Value
                ws.Cells(Summary_Table_Row, "J").Value = Yearly_Change
                ws.Cells(Summary_Table_Row, "K").Value = Percent_Change
                ws.Cells(Summary_Table_Row, "L").Value = Volume
                
                If Yearly_Change > 0 Then
                    ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
                End If
        
                If Percent_Change > 0 Then
                    ws.Cells(Summary_Table_Row, "K").Interior.ColorIndex = 4
                Else
                    ws.Cells(Summary_Table_Row, "K").Interior.ColorIndex = 3
                End If
        
        
        
                Summary_Table_Row = Summary_Table_Row + 1
                Open_Price_Pointer = i + 1
                
                Volume = 0
            
                    
            End If
       
        Next i

            
            ws.Cells(2, "Q") = greatest_increase
            ws.Cells(2, "P") = gi_ticker
            ws.Cells(2, "Q").NumberFormat = "0.00%"
            
            ws.Cells(3, "Q") = greatest_decrease
            ws.Cells(3, "P") = gd_ticker
            ws.Cells(3, "Q").NumberFormat = "0.00%"
            
            ws.Cells(4, "Q") = greatest_volume
            ws.Cells(4, "P") = gv_ticker
            
            
Next ws

MsgBox ("Complete")

End Sub
