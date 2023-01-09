
Sub stockprices():

' ensure that script loops through each worksheet

For Each ws In Worksheets

' Name columns

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"



' set variables
        ' capture the last row of the data the data set
        ' capture the data needed to complete the challenge
        ' set range where I am  putting the data
        
Dim last_row As Long
Dim ticker_symbol As String
Dim summary_table_row As Integer
Dim opeing_price As Double
Dim closing_price As Double
Dim volume_total As Double

summary_table_row = 2
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
volume_total = 0
opening_price = ws.Cells(2, 3).Value

' Loop through each trading day

For i = 2 To last_row

' Check if we are within the same ticker symbol on the next immediate row

      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

' if not then set ticker variable to the preceding ticker and closing_price to the closing price on the last day
      
      ticker_symbol = ws.Cells(i, 1).Value
      closing_price = ws.Cells(i, 6).Value
      
      

    
      
' 1) add last trading day volume to volume variable
' 2) run calculations and place values in correct ranges

      volume_total = volume_total + ws.Cells(i, 7).Value
      
      ws.Range("I" & summary_table_row).Value = ticker_symbol
      
      ws.Range("J" & summary_table_row).Value = closing_price - opening_price
      
    ws.Range("K" & summary_table_row).Value = (closing_price - opening_price) / opening_price
      
      ws.Range("L" & summary_table_row).Value = volume_total
      
' move to next row in summary table
      
      summary_table_row = summary_table_row + 1
      
' reset volume total to 0 and reset opeing price the opening price on the first day for the next ticker
      
      volume_total = 0
      
      opening_price = ws.Cells(i + 1, 3).Value
      
      
      Else

' Grab volume and update volume variable
      
      volume_total = volume_total + ws.Cells(i, 7).Value
      
      End If
      
    Next i
    
' format the percentage change row to percent
    
    ws.Range("K2:K" & last_row).NumberFormat = "0.00%"
    
Next ws


End Sub
