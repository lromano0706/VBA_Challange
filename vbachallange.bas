Attribute VB_Name = "Module1"
Sub stocks()





'* You should also have conditional formatting that will highlight positive change in green and negative change in red.



For Each ws In Worksheets
    ws.Activate
    Dim ticker As String
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    openvalue = Cells(2, 3).Value
    Dim percentagechange As Double
    SummaryRow = 2
    ws.Range("i1").Value = "Ticker"
    ws.Range("j1").Value = "Volume"
    ws.Range("k1").Value = "Yearly Change"
    ws.Range("l1").Value = "Percentage Change"
    Range("l:l").NumberFormat = "0.00%"
    
    
    For i = 2 To LastRow

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            Volume = Volume + ws.Cells(i, 7).Value
            closevalue = ws.Cells(i, 6).Value
            ws.Cells(SummaryRow, 9).Value = ticker
            ws.Cells(SummaryRow, 10).Value = Volume
            yearlychange = closevalue - openvalue
            ws.Cells(SummaryRow, 11).Value = yearlychange
            percentagechange = yearlychange / openvalue
            ws.Cells(SummaryRow, 12).Value = percentagechange
            Volume = 0
            SummaryRow = SummaryRow + 1
            openvalue = Cells(i + 1, 3)
           
      
        Else
            Volume = Volume + ws.Cells(i, 7).Value
        

        End If

            
     Next i
        

    
        
Next ws

   


End Sub
