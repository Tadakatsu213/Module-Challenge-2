Sub Multiple_year_stock_data()
    Dim ws As Worksheet
    Dim ticker As String
    Dim lastRow As Long
    Dim summaryTableRow As Integer
    Dim openPrice As Double
    Dim closePrice As Double
    Dim totalVolume As Double
    Dim change As Double
    Dim percentChange As Double

For Each ws In Worksheets
    ws.Activate
        
    summaryTableRow = 2
    openPrice = ws.Cells(2, 3).Value
    totalVolume = 0
        
    ' Find the last row with data
lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
    
For i = 2 To lastRow
    totalVolume = totalVolume + ws.Cells(i, 7).Value
            
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
               
    ticker = ws.Cells(i, 1).Value
    closePrice = ws.Cells(i, 6).Value
                
' Calculating the change and percent change
    change = closePrice - openPrice
    If openPrice <> 0 Then
       percentChange = (change / openPrice) * 100
        Else
        percentChange = 0
        End If
                
'the data to the summary table
     ws.Cells(summaryTableRow, 9).Value = ticker
      ws.Cells(summaryTableRow, 10).Value = change
     ws.Cells(summaryTableRow, 11).Value = percentChange
       ws.Cells(summaryTableRow, 12).Value = totalVolume
       
'condtional formating
       If quarterlyChange > 0 Then
        ws.Cells(summaryTableRow, 10).Interior.Color = vbGreen
        ElseIf quarterlyChange < 0 Then
        ws.Cells(summaryTableRow, 10).Interior.Color = vbRed
            End If
                
             
    summaryTableRow = summaryTableRow + 1
                
        If i < lastRow Then
        openPrice = ws.Cells(i + 1, 3).Value
            totalVolume = 0 '
End If
    End If
        Next i
    Next ws
End Sub

