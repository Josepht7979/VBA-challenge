Sub Multiple_year_stock_analysis():

For Each ws In Worksheets
   ws.Range("I1").Value = "Ticker"
  ws.Range("J1").Value = "Yearly Change"
  ws.Range("K1").Value = "Percent Change"
  ws.Range("L1").Value = "Total Stock Volume"
  ws.Range("O2").Value = "Greatest % Increase"
  ws.Range("O3").Value = "Greatest % Decrease"
  ws.Range("O4").Value = "GreatestTotal Volume"
  ws.Range("P1").Value = "Ticker"
  ws.Range("Q1").Value = "Value"
  
Dim rowCount, PercentCount As Long
rowCount = 0
 Dim tickerResult, tickerMin, tickerMax, max_vol_tick As String
 Dim totalStock, yearlyChange, percentChange, openYearly, closeYearly  As Double
  totalStock = 0
  yearlyChange = 0
  percentChange = 0
  openYearly = 0
  closeYearly = 0
 
   Dim Resultrow  As Integer
  Resultrow = 2
  
  ' Record stock open value
  Dim Flag As Boolean
    '----------------------------------------
  Dim min, max, max_vol As Double
  
'Go through until last row
  rowCount = rowCount + ws.Cells(Rows.Count, "A").End(xlUp).Row
    
  For i = 2 To rowCount
    
    If (ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value) Then
      tickerResult = ws.Cells(i, 1).Value
 
  totalStock = totalStock + ws.Cells(i, 7).Value
      closeYearly = ws.Cells(i, 6).Value
     yearlyChange = closeYearly - openYearly
         
      
'  Printing results of Ticker I , Yearly change J, percentage change K , total stock Volume L
      
       ' Print the Ticker Result
        ws.Range("I" & Resultrow).Value = tickerResult
      
       ' Print Total Stock Volume
        ws.Range("L" & Resultrow).Value = totalStock
      
       'Print Yearly change
        ws.Range("J" & Resultrow).Value = yearlyChange
        
        
         'Colour Yearly Change Column
                    If (yearlyChange > 0) Then
       ws.Range("J" & Resultrow).Interior.ColorIndex = 4
                   Else
                     ws.Range("J" & Resultrow).Interior.ColorIndex = 3
                         End If
         
'Print percent change
            If (yearlyChange <> 0) Then
               percentChange = FormatPercent((percentChange + (closeYearly - openYearly) / openYearly))
            Else
               percentChange = 0
            End If
 ws.Range("K" & Resultrow).Value = percentChange
   
'---------------------------------------------------------------------------------------
     Resultrow = Resultrow + 1
      
      
        totalStock = 0
        yearlyChange = 0
        percentChange = 0
        openYearly = 0
        closeYearly = 0
        Flag = False
Else
       
' Add total stock
      totalStock = totalStock + ws.Cells(i, 7).Value

 'Open Yearly
       If (Flag = False) And (ws.Cells(i, 3).Value <> 0) Then
             openYearly = ws.Cells(i, 3).Value
             Flag = True
        End If
    End If
  Next i


'Go through Percent column to get the Greatest% increase, Greatest% decrease
  PercentCount = 0
  min = 0
  max = 0
  max_vol = 0
  
  PercentCount = PercentCount + ws.Cells(Rows.Count, "K").End(xlUp).Row
         
        For j = 2 To PercentCount
             ' Check for min
                 If (ws.Cells(j, 11).Value < min) Then
                         min = (ws.Cells(j, 11).Value)
                         tickerMin = (ws.Cells(j, 9).Value)
                         
               ' Greatest percentage decrease
                   ElseIf (ws.Cells(j, 11).Value > max) Then
                             max = (ws.Cells(j, 11).Value)
                             tickerMax = (ws.Cells(j, 9).Value)
                    End If
                         
                 ' Greatest total Volume
                                 If (ws.Cells(j, 12).Value > max_vol) Then
                                    max_vol = (ws.Cells(j, 12).Value)
                                    max_vol_tick = (ws.Cells(j, 9).Value)
                                
                                End If
                 Next j
          
     'Print the final output
        ws.Range("P" & 2).Value = tickerMax
        ws.Range("Q" & 2).Value = FormatPercent(max)
        ws.Range("P" & 3).Value = tickerMin
        ws.Range("Q" & 3).Value = FormatPercent(min)
        ws.Range("P" & 4).Value = max_vol_tick
        ws.Range("Q" & 4).Value = max_vol
        
   Next ws
End Sub

