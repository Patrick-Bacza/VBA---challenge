
Sub conditional_formatting():

' ensure it runs for each worksheet

For Each ws In Worksheets

' calculate last row in each worksheet

Dim last_row As Long

last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row

' Use for loop/if statement to apply conditional formatting to columns

For i = 2 To last_row

' First if statement for yearly change

    If ws.Cells(i, 10) > 0 Then

    ws.Cells(i, 10).Interior.ColorIndex = 4

    ElseIf ws.Cells(i, 10) < 0 Then

    ws.Cells(i, 10).Interior.ColorIndex = 3

    Else
    
    ws.Cells(i, 10).Interior.ColorIndex = 2
    
    End If
    
' Second if statement for percentage change column
    
    If ws.Cells(i, 11) > 0 Then

    ws.Cells(i, 11).Interior.ColorIndex = 4

    ElseIf ws.Cells(i, 11) < 0 Then

    ws.Cells(i, 11).Interior.ColorIndex = 3

    Else
    
    ws.Cells(i, 11).Interior.ColorIndex = 2
    
    End If
    
Next i

Next ws


End Sub

