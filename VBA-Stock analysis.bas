Attribute VB_Name = "Module1"
 Sub Stock_Returns()
    Dim total_stock As Double
    Dim change As Double
    Dim start As Long
    Dim i As Long
    Dim j As Integer
    Dim row_count As Double
    Dim percent_change As Double
    Dim days As Integer
    Dim daily_change As Double
    Dim ave_change As Double
    Dim ws As Worksheet
    
For Each ws In ThisWorkbook.Worksheets
    total_stock = 0
    change = 0
    daily_change = 0
    start = 2
    i = 0
    j = 0
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    row_count = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To row_count
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            total_stock = total_stock + Cells(i, 7).Value
            
            If total_stock = 0 Then
                ws.Cells(2 + j, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(2 + j, 10).Value = 0
                ws.Cells(2 + j, 11).Value = "%" & 0
                ws.Cells(2 + j, 12).Value = 0
                
            Else
                If ws.Cells(start, 3).Value = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                        End If
                    Exit For
                    Next find_value
                End If
                
                'Calculate yearly change
                change = (ws.Cells(i, 6).Value - ws.Cells(start, 3).Value)
                                
                'Calculate percent change -
                percent_change = Round((change / ws.Cells(start, 3).Value * 100), 2)
                
                'Move to the next ticker
                start = i + 1
                
                    ws.Cells(2 + j, 9).Value = ws.Cells(i, 1).Value
                    ws.Cells(2 + j, 10).Value = change
                    ws.Cells(2 + j, 11).Value = "%" & percent_change
                    ws.Cells(2 + j, 12).Value = total_stock
                                
                'Color the Cells
                Select Case change
                    Case Is > 0
                        ws.Cells(2 + j, 10).Interior.ColorIndex = 4
                    Case Is < 0
                        ws.Cells(2 + j, 10).Interior.ColorIndex = 3
                End Select
            
                'Increase J
                j = j + 1
            'Now that cells are filled and colored we can move to next item
            End If
            
        total_stock = 0
        change = 0
        days = 0
        daily_change = 0
        End If
    Next i
Next ws
End Sub



