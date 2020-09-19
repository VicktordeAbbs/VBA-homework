
Option Explicit
Sub TickerFirst()
    
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
       
        Dim myTicker As String
        Dim Stock_TotalVol As Double
        Dim Yearly_change As Double
        Dim Percentage_Change As Variant
            
        Stock_TotalVol = 0
        
        Dim Startrow As Double
        Dim Lastrow As Double
        Dim i As Double
        Dim TickerCount As Double
        Dim Opening_Price As Double
        Dim Closing_Price As Double
        Dim Range_1 As Variant
        Dim Range_2 As Variant
        Dim Range_3 As Variant
            
        
        ws.Cells(1, 9).Value = "Tickers"
        ws.Cells(1, 10).Value = "Yearly change"
        ws.Cells(1, 11).Value = "% change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
       
            
        Startrow = 2
        Lastrow = ws.Range("A" & 2).End(xlDown).Row
        
        ' Loop through columnA
        For i = 2 To Lastrow
      
            'Opening price_Initial
            Opening_Price = ws.Cells(2, 3).Value
            
            ' Check if we are still within the same Ticker Group, if it is not...
                    If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                        myTicker = ws.Cells(i, 1).Value
                        
                        'Print Ticker name to summary
                        ws.Range("I" & Startrow).Value = myTicker
                                    
                        'Add to Total Stock Volume
                        Stock_TotalVol = Stock_TotalVol + ws.Cells(i, 7).Value
                        
                        'Print stock total volume to summary
                        ws.Range("L" & Startrow).Value = Stock_TotalVol
                        
                        
                        'Reset Stock_TotalVol
                        Stock_TotalVol = 0
                            
                                                
                       'Opening and closing price prices
                       
                       Closing_Price = ws.Cells(i, 6).Value
                       Opening_Price = ws.Cells(i - TickerCount, 3).Value
                       Yearly_change = Closing_Price - Opening_Price
           
                            'Percentage_Change
                            If Closing_Price <> 0 And Opening_Price <> 0 Then
                            
                            Percentage_Change = (Closing_Price - Opening_Price) / Opening_Price
                            'Percentage_Change = (ws.Cells(i, 6).Value - ws.Cells(i - TickerCount, 3).Value) / (ws.Cells(i - TickerCount, 3).Value)
                            Else: Percentage_Change = 0
                            End If
                
                            
                    'Print Yearly_change to summary
                     ws.Range("J" & Startrow).Value = Yearly_change
            
                            'Format yearly change color
                            If Yearly_change > 0 Then
                                ws.Range("J" & Startrow).Interior.ColorIndex = 4
                                Else:
                                ws.Range("J" & Startrow).Interior.ColorIndex = 3
                            End If
            
                    'Print Percentage_Change
                    ws.Range("K" & Startrow).Value = Format(Percentage_Change, "Percent")
                                 
                    Startrow = Startrow + 1
                    'Reset Ticker counter
                    TickerCount = 0
 
           
            Else
            'CountTicker
                        
            TickerCount = TickerCount + ws.Cells(i, 1).Count
                               
            'Add to StockTotalVol
            Stock_TotalVol = Stock_TotalVol + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
        'Challenges section
        'Columns and rows labels
                
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
        '% change column range
        Range_1 = ws.Range("K2", "K" & Cells(Rows.Count, 11).End(xlUp).Row)
        'Total stock vol column range
        Range_2 = ws.Range("L2", "L" & Cells(Rows.Count, 12).End(xlUp).Row)
        Range_3 = ws.Range("I2", "I" & Cells(Rows.Count, 9).End(xlUp).Row)
        
        ' Obtain and print Greatest % increase, greatest total volume and greatest %decrease
        
        ws.Cells(2, 16).Value = Application.WorksheetFunction.Max(Range_1)
        ws.Cells(3, 16).Value = Application.WorksheetFunction.Min(Range_1)
        ws.Cells(4, 16).Value = Application.WorksheetFunction.Max(Range_2)
                
        'Print Ticker corresponding to Greatest % increase, greatest total volume and greatest %decrease

        ws.Cells(2, 15).Value = WorksheetFunction.Index(Range_3, WorksheetFunction.Match(WorksheetFunction.Max(Range_1), Range_1, 0))
        ws.Cells(3, 15).Value = WorksheetFunction.Index(Range_3, WorksheetFunction.Match(WorksheetFunction.Min(Range_1), Range_1, 0))
        ws.Cells(4, 15).Value = WorksheetFunction.Index(Range_3, WorksheetFunction.Match(WorksheetFunction.Max(Range_2), Range_2, 0))
        
        ws.Range("I:R").Columns.AutoFit
                        
    Next ws
        
End Sub
        
   
