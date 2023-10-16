Attribute VB_Name = "Module6"
'Module 2 - VBA challenge - Assignment
'Person = Vy Nguyen
'Date: 16/10/2023

Sub StockChecker()

    For Each ws In Worksheets
    
        'Define variables
         Dim NumRow As Long
         Dim TickerCount As Long
         Dim StartRow As Long
         Dim OpenPrice As Double
         Dim ClosePrice As Double
         Dim YearlyChange As Double
         Dim PercentChange As Double
         Dim TotalStock As LongLong
         Dim NumTicker As Long
         Dim GreatestIncrease As Double
         Dim GreatestDecrease As Double
         Dim GreatestVolume As LongLong
         
        'Naming columns and rows
         ws.Range("I1").Value = "Ticker"
         ws.Range("J1").Value = "Yearly Change"
         ws.Range("K1").Value = "Percent Change"
         ws.Range("L1").Value = "Total Stock Value"
         ws.Range("L1").Value = "Total Stock Value"
         ws.Range("P1").Value = "Ticker"
         ws.Range("Q1").Value = "Value"
         ws.Range("O2").Value = "Greatest%Increase"
         ws.Range("O3").Value = "Greatest%Decrease"
         ws.Range("O4").Value = "Greatest Total Volume"
               
        'Find the number of rows in the data
         NumRow = ws.Cells(Rows.Count, 1).End(xlUp).row
        
        'Create a script that loops through all the stocks for one year
         TickerCount = 2
         StartRow = 2
         
         For i = 2 To NumRow
             
             If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
             
                 'Output the sticker symbol to the "Ticker" column
                 ws.Range("I" & TickerCount).Value = ws.Cells(i, 1).Value
                 
                 'Output yearly change (final close price - beginning open price) to "Yearly Change" column
                 ClosePrice = ws.Cells(i, 6).Value
                 OpenPrice = ws.Cells(StartRow, 3).Value
                 YearlyChange = ClosePrice - OpenPrice
                 ws.Range("J" & TickerCount).Value = YearlyChange
                 
                 'Output percentage change (yearly change / beginning open price) to "Percentage Change" column
                 PercentChange = YearlyChange / OpenPrice
                 ws.Range("K" & TickerCount).Value = PercentChange
                 ws.Range("K" & TickerCount).NumberFormat = "0.00%"
                 
                 'Calculate Total Stock volume and output the result to "Total Stock Value" column
                 ws.Range("L" & TickerCount).Value = WorksheetFunction.Sum(Range(ws.Cells(StartRow, 7), ws.Cells(i, 7)))
                 
                 'Formatting cells in "Yearly Change" Column base on pos and neg value
                     If ws.Range("J" & TickerCount).Value > 0 Then
                         ws.Range("J" & TickerCount).Interior.ColorIndex = 4
                     Else
                         ws.Range("J" & TickerCount).Interior.ColorIndex = 3
                     End If
                     
                 'Formatting cells in "Percent Change" Column base on pos and neg value
                     If ws.Range("K" & TickerCount).Value > 0 Then
                         ws.Range("K" & TickerCount).Interior.ColorIndex = 4
                     Else
                         ws.Range("K" & TickerCount).Interior.ColorIndex = 3
                     End If
                              
                 'Set new starting data point for the next loop
                 TickerCount = TickerCount + 1
                 StartRow = i + 1
                 
             End If
             
          Next i
         
         '''BONUS'''
         
         'Count the number of rows in the "Ticker" colum
          NumTicker = ws.Cells(Rows.Count, 9).End(xlUp).row
         
         'Return stock with "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
          GreatestIncrease = ws.Range("K2").Value
          GreatestDecrease = ws.Range("K2").Value
          GreatestVolume = ws.Range("L2").Value
         
          For i = 2 To NumTicker
         
             'Greatest % increase
              If ws.Range("K" & i) >= GreatestIncrease Then
                 GreatestIncrease = ws.Range("K" & i)
                 ws.Range("P2").Value = ws.Range("I" & i).Value
                 ws.Range("Q2").Value = ws.Range("K" & i).Value
                 ws.Range("Q2").NumberFormat = "0.00%"
              Else
                 GreatestIncrease = GreatestIncrease
              End If
             
             'Greatest % decrease
              If ws.Range("K" & i) <= GreatestDecrease Then
                 GreatestDecrease = ws.Range("K" & i)
                 ws.Range("P3").Value = ws.Range("I" & i).Value
                 ws.Range("Q3").Value = ws.Range("K" & i).Value
                 ws.Range("Q3").NumberFormat = "0.00%"
              Else
                 GreatestDecrease = GreatestDecrease
              End If
         
             'Greatest total volume
              If ws.Range("L" & i) >= GreatestVolume Then
                 GreatestVolume = ws.Range("L" & i)
                 ws.Range("P4").Value = ws.Range("I" & i).Value
                 ws.Range("Q4").Value = ws.Range("L" & i).Value
              Else
                 GreatestVolume = GreatestVolume
              End If
                    
         Next i
         
    Next ws
    
End Sub


