
Sub greatest():

' Ensure script loops through each worksheet

For Each ws In Worksheets

' Name columns

ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % decrease"
ws.Range("O4").Value = "Greatest total volume"

' Create variables needed

Dim greatest_i_ticker As String
Dim greatest_d_ticker As String
Dim greatest_v_ticker As String
Dim max_increase As Double
Dim max_decrease As Double
Dim max_volume As Double

'Capture the last row into a variable and set Max/min variables to the maximum and minimum values being asked for

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

max_increase = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_row))
max_decrease = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_row))
max_volume = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_row))

' For loop is used to grab the ticker symbol
' loop through the correct columns and find the max/min and grab the ticker symbol for both


 For i = 2 To last_row

     If ws.Cells(i, 11).Value = max_increase Then
     
     greatest_i_ticker = ws.Cells(i, 9).Value
     
     ElseIf ws.Cells(i, 11).Value = max_decrease Then
     
     greatest_d_ticker = ws.Cells(i, 9).Value
     
     End If

' Used a different if for volume. If the ticker with the greatest increase/decrease also had the most volume it would not have been captured if I used one if statement for all three values

     
     If ws.Cells(i, 12).Value = max_volume Then
     
     greatest_v_ticker = ws.Cells(i, 9).Value
     
     End If
     
   Next i
     
 ' add all values to their respective cells
 
     ws.Range("P2").Value = greatest_i_ticker
     ws.Range("Q2").Value = max_increase
     ws.Range("P3").Value = greatest_d_ticker
     ws.Range("Q3").Value = max_decrease
     ws.Range("P4").Value = greatest_v_ticker
     ws.Range("Q4").Value = max_volume
     
' formatted percentage change column as percent
     ws.Range("Q2:Q3").NumberFormat = "0.00%"
     
Next ws

End Sub

