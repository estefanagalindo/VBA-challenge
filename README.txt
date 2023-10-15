Sub stockHard()
    Dim currentName As String
    Dim nextName As String
    Dim totalSV As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim i As Long
    Dim lastRow As Long
    Dim ws As Worksheet
   
   
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Charge"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volumn"
       
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greates Total Volume"
       
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
       
        totalSV = 0
        groupNo = 1
       
        openPrice = ws.Cells(2, 3).Value
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
       
       
        For i = 2 To lastRow
           
            ws.Range("K" & i).NumberFormat = "0.00%"
            currentName = ws.Cells(i, 1).Value
            nextName = ws.Cells(i + 1, 1).Value
           
            If nextName = currentName Then
                totalSV = totalSV + ws.Cells(i, 7).Value
            Else
                totalSV = totalSV + ws.Cells(i, 7).Value
                closePrice = ws.Cells(i, 6).Value
               
                YrChange = closePrice - openPrice
                PctChange = YrChange / openPrice
               
                ws.Cells(groupNo + 1, 9).Value = currentName
                ws.Cells(groupNo + 1, 10).Value = YrChange
                ws.Cells(groupNo + 1, 11).Value = PctChange
                ws.Cells(groupNo + 1, 12).Value = totalSV
               
                totalSV = 0
                openPrice = ws.Cells(i + 1, 3).Value
                groupNo = groupNo + 1
               
               
               
            End If
           
        Next i
       
        ws.Range("Q2") = WorksheetFunction.Max(ws.Range("K2:K" & lastRow))
        ws.Range("Q3") = WorksheetFunction.Min(ws.Range("K2:K" & lastRow))
        ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & lastRow))
       
        max_change_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        min_change_index = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastRow)), ws.Range("K2:K" & lastRow), 0)
        max_volume_index = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastRow)), ws.Range("L2:L" & lastRow), 0)

        ws.Range("P2") = ws.Cells(max_change_index + 1, 9)
        ws.Range("P3") = ws.Cells(min_change_index + 1, 9)
        ws.Range("P4") = ws.Cells(max_volume_index + 1, 9)
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"
       
    Next
   
End Sub
