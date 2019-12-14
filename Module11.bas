Attribute VB_Name = "Module11"
Sub StockData():

For Each ws In Worksheets:


Dim table As Integer
table = 2
openrow = 2

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"


lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow:
   
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        ticker = ws.Cells(i, 1).Value
        opencost = ws.Cells(openrow, 3).Value
        If opencost = 0 Then
            opencost = 1
        End If
        
        closecost = ws.Cells(i, 6).Value
        yearchg = closecost - opencost
        percentchg = yearchg / opencost
        volumetotal = volumetotal + ws.Cells(i, 7).Value
        
        
        ws.Range("I" & table).Value = ticker
        ws.Range("J" & table).Value = yearchg
        ws.Range("K" & table).Value = percentchg
        ws.Range("L" & table).Value = volumetotal
         ws.Range("K" & table).NumberFormat = "0.00%"
        If ws.Range("J" & table).Value > 0 Then
        ws.Range("J" & table).Interior.ColorIndex = 4
        ElseIf ws.Range("J" & table).Value < 0 Then
        ws.Range("J" & table).Interior.ColorIndex = 3
    
    End If
        
        
        
        openrow = (i + 1)
        table = table + 1
        volumetotal = 0
    Else
        volumetotal = volumetotal + ws.Cells(i, 7).Value
    
    End If
    
    
    
    
    
    
Next i


Next ws
    




End Sub
