
Sub Stockmarket()

    For Each ws In Worksheets
    ws.Activate


        'Header
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change"
        Range("M1").Value = "Total Stock Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        
        'Set Variables & Type
        
        Dim Ticker_Symbol As String
        Dim Total_Stock_Volume As Double
        Dim Last_row As Long
        Dim Open_Year_Row As Long
        Dim Open_Year_Price As Double
        Dim End_Year_Price As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double
        Dim Table_Row As Integer
        
        'Initial Values
        Table_Row = 2
        Open_Year_Row = 2
        Total_Stock_Volume = 0
        Last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'For Loop
        For i = 2 To Last_row

            'Sum up the stock volume for each ticker
            If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
            Else
                'List Ticker
                Ticker_Symbol = Cells(i, 1).Value
                Range("J" & Table_Row).Value = Ticker_Symbol
                
                'Find yearly change, percent change for each ticker
                End_Year_Price = Cells(i, 6).Value
                Open_Year_Price = Cells(Open_Year_Row, 3).Value
                Yearly_Change = (End_Year_Price - Open_Year_Price)
                Range("K" & Table_Row).Value = Yearly_Change
                
                'Conditional for avoiding error for having 0 in divisor
                If Open_Year_Price = 0 Then
                Percent_Change = 0
                Range("L" & Table_Row).Value = Percent_Change
                Else
                
                Percent_Change = (Yearly_Change / Open_Year_Price)
                Range("L" & Table_Row).Value = Percent_Change
                End If
            
                Range("L" & Table_Row).NumberFormat = "0.00%"
                                      
                Range("M" & Table_Row).Value = Total_Stock_Volume + Cells(i, 7).Value
                
    
                               
                If Range("K" & Table_Row).Value > 0 Then
            'green
                    Range("K" & Table_Row).Interior.ColorIndex = 4
                Else
             'red
                    Range("K" & Table_Row).Interior.ColorIndex = 3
                End If
                               
               'For next loop
                Table_Row = Table_Row + 1
                Total_Stock_Volume = 0
                Open_Year_Row = i + 1
      
            End If
            
        Next i


    'Variables Bonus
        Dim lr As Integer
        Dim Most_Inc As Double
         Dim Most_Dec As Double
        Dim Most_Total_Vol As Double
        Most_Inc = 0
        Most_Dec = 0
        Most_Total_Vol = 0
        lr = ws.Cells(Rows.Count, 11).End(xlUp).Row

        For j = 2 To lr
                'Greatest Increase    
            If Range("L" & j).Value > Most_Inc Then
                Range("P2").Value = Range("J" & j).Value
                Most_Inc = Range("L" & j).Value
                Range("Q2").Value = Most_Inc
                Range("Q2").NumberFormat = "0.00%"
                
               'Greatest Decrease           
            ElseIf Range("L" & j).Value < Most_Dec Then
                Range("P3").Value = Range("J" & j).Value
                Most_Dec = Range("L" & j).Value
                Range("Q3").Value = Most_Dec
                Range("Q3").NumberFormat = "0.00%"

            End If
            'Greatest stock volume
            If Range("M" & j).Value > Most_Total_Vol Then
                Range("P4").Value = Range("J" & j).Value
                Most_Total_Vol = Range("M" & j).Value
                Range("Q4").Value = Most_Total_Vol
                
            
            End If
  
            
        Next j

    Next ws

End Sub

