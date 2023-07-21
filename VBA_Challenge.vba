Attribute VB_Name = "Module1"
'Thank you to my tutor Limei Hou for helping me apply these concepts

Sub StockData():
    Dim currentName As String
    Dim nextName As String
    Dim totalSV As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim i As Long
    Dim lastRow As Long
    Dim curSheet As Worksheet
    
    For Each curSheet In ActiveWorkbook.Worksheets
        curSheet.Activate
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        
        totalSV = 0
        groupNo = 1
        
        openPrice = Cells(2, 3).Value
        lastRow = Cells(curSheet.Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To lastRow
        
            currentName = Cells(i, 1).Value
            nextName = Cells(i + 1, 1).Value
            
            If nextName = currentName Then
                totalSV = totalSV + Cells(i, 7).Value
            Else
                totalSV = totalSV + Cells(i, 7).Value
                closePrice = Cells(i, 6).Value
                YrChange = closePrice - openPrice
                PctChange = YrChange / openPrice
           
                curSheet.Cells(groupNo + 1, 9).Value = currentName
                curSheet.Cells(groupNo + 1, 10).Value = YrChange
                curSheet.Cells(groupNo + 1, 11).Value = PctChange
                curSheet.Cells(groupNo + 1, 12).Value = totalSV
                
                totalSV = 0
                openPrice = Cells(i + 1, 3).Value
                groupNo = groupNo + 1
                
            End If
                
        Next i
    Next
End Sub
