Attribute VB_Name = "Module1"
Sub alphabetical_testing()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call RunCode
    Next
    Application.ScreenUpdating = True
End Sub
Sub RunCode()

    Dim ticker, ticker_mayor, ticker_menor As String
    Dim totalstockvol, mayor_stockvol As Double
    Dim summary_table_row As Integer
    Dim opening_price As Double
    Dim closing_price As Double
    Dim yearly_change As Double
    Dim mayor, menor As Double

    
    Range("J1").FormulaR1C1 = "Ticker"
    Range("K1").FormulaR1C1 = "Yearly Change"
    Range("L1").FormulaR1C1 = "Percentage Change"
    Range("M1").FormulaR1C1 = "Total Stock Volume"
    Range("P1").FormulaR1C1 = "Ticker"
    Range("Q1").FormulaR1C1 = "Value"
    
    Range("J1:Q1").Font.Bold = True
    Range("J1:Q1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range("O2").FormulaR1C1 = "Greatest % Increase"
    Range("O3").FormulaR1C1 = "Greatest % Decrease"
    Range("O4").FormulaR1C1 = "Greatest Total Volume"
    
    Range("O2:O4").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
    
    
    summary_table_row = 2
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    
    For i = 2 To lastRow
    
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            opening_price = Cells(i, 3).Value
        
        ElseIf Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            ticker = Cells(i, 1).Value
            totalstockvol = totalstockvol + Cells(i, 7).Value
            Cells(summary_table_row, 10).Value = ticker
            Cells(summary_table_row, 13).Value = totalstockvol
            
            closing_price = Cells(i, 6).Value
            yearly_change = closing_price - opening_price
            Cells(summary_table_row, 11).Value = yearly_change
            
            If yearly_change >= 0 Then
                Cells(summary_table_row, 11).Interior.ColorIndex = 4
            
            Else
                Cells(summary_table_row, 11).Interior.ColorIndex = 3
            
            End If
            
            If opening_price = 0 Then
                opening_price = 0.01
                percentage_change = yearly_change / opening_price
            Else
                percentage_change = yearly_change / opening_price
            End If
                
            Cells(summary_table_row, 12).Value = percentage_change
            Cells(summary_table_row, 12).Select
            Selection.Style = "Percent"
               
            summary_table_row = summary_table_row + 1
            
            totalstockvol = 0
            
        Else
            totalstockvol = totalstockvol + Cells(i, 7).Value
        
        End If
        
    Next i

    lastRow2 = Cells(Rows.Count, "L").End(xlUp).Row
    summary_table_row2 = 2
    mayor = Cells(2, 12).Value
    menor = Cells(2, 12).Value
    mayor_stockvol = Cells(2, 13).Value
    
        
    For j = 2 To lastRow2
    
        If Cells(j, 12).Value > mayor Then
            mayor = Cells(j, 12).Value
            ticker_mayor = Cells(j, 10).Value
                                   
        End If
         
        If Cells(j, 12).Value < menor Then
            menor = Cells(j, 12).Value
            ticker_menor = Cells(j, 10).Value
                                   
        End If
        
        If Cells(j, 13).Value > mayor_stockvol Then
            mayor_stockvol = Cells(j, 13).Value
            ticker_mayor_stockvol = Cells(j, 10).Value
                                   
        End If
        
        
    Next j
    
    Cells(summary_table_row2, 17).Value = mayor
    Cells(summary_table_row2, 17).Select
    Selection.Style = "Percent"
    
    Cells(summary_table_row2, 16).Value = ticker_mayor
    
    
    summary_table_row2 = summary_table_row2 + 1
    
    Cells(summary_table_row2, 17).Value = menor
    Cells(summary_table_row2, 17).Select
    Selection.Style = "Percent"
    
    Cells(summary_table_row2, 16).Value = ticker_menor
    
    
    summary_table_row2 = summary_table_row2 + 1
    
    Cells(summary_table_row2, 17).Value = mayor_stockvol
    
    Cells(summary_table_row2, 16).Value = ticker_mayor_stockvol
    
    Columns("J:Q").EntireColumn.AutoFit
    Range("J1:M1").AutoFilter
    
End Sub



