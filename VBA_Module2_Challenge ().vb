Sub VBA_Module2_Challenge ()


Dim Ticker_Name As String
Dim Volume As Double
Dim Open_Price As Double
Dim Close_Price As Double
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim Summary_Table As Integer
Dim Rng As Range



For Each ws In ThisWorkbook.Worksheets
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    
    lastRow2 = Range("A" & Rows.Count).End(xlDown).Row
    
    Summary_Table_Row = 2

    
        For i = 2 To lastRow2
             Ticker = ws.Cells(i, 1).Value
             Ticker_Next = ws.Cells(i + 1, 1).Value
             If Ticker <> Ticker_Next Then
             Ticker_Name = ws.Cells(i, 1).Value
             Volume = ws.Cells(i, 7).Value

            Open_Price = ws.Cells(i, 3).Value
            Close_Price = ws.Cells(i, 6).Value
            Yearly_Change = (Close_Price - Open_Price)
            
                If (Open_Price = 0) Then
                    Percent_Change = 0
                Else
                    Percent_Change = Yearly_Change / Open_Price
                End If
            
            

           
            ws.Cells(Summary_Table_Row, 9).Value = Ticker_Name
            ws.Cells(Summary_Table_Row, 10).Value = Yearly_Change
            ws.Cells(Summary_Table_Row, 11).Value = Percent_Change
            ws.Cells(Summary_Table_Row, 12).Value = Volume
            Summary_Table_Row = Summary_Table_Row + 1

            Volume = 0
        
            End If
       Next i
       
       lastRow = Range("J" & Rows.Count).End(xlDown).Row
       For j = 2 To lastRow
        
        
        If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 10
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
        End If

    Next j
    
ws.Columns("K").NumberFormat = "0.00%"






Next ws


End Sub