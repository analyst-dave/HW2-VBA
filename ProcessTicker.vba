Attribute VB_Name = "Module1"
Sub ProcessTicker()

Dim WS As Worksheet

For Each WS In ActiveWorkbook.Worksheets
    
    WS.Activate
    
    Dim startYearD As Double
    Dim endYearD As Double
    Dim yearChangeD As Double
    Dim sumLL As LongLong
    Dim counter1 As Integer
    Dim counter2 As Integer
    
    ' starts from 1st row of record
    counter1 = 2
    counter2 = 0

    ' startYearD = Cells(2, 3).Value
    Cells(1, 9).Value = "<Ticker Symbol>"
    Cells(1, 10).Value = "<Yearly Change>"
    Cells(1, 11).Value = "<Percentage Change>"
    Cells(1, 12).Value = "<Total Volume>"
    
    ' iterate until the last row
    For i = 2 To Cells(Rows.Count, 6).End(xlUp).Row

        
        ' if next symbol is new
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
        
            ' extract values from cells
            startYearD = Cells(i - counter2, 3).Value
            ' endYearD = Cells(i, 6).Value
            endYearD = Cells(i, 6).Value
            sumLL = sumLL + Cells(i, 7).Value
            
            temp1 = endYearD - startYearD
            If startYearD <> 0 Then
                yearChangeD = temp1 / startYearD
            End If
            
            ' assign values in cells
            Cells(counter1, 9).Value = Cells(i, 1).Value
            Cells(counter1, 10).Value = Format(temp1, "$#,###.##")
            Cells(counter1, 11).Value = Format(yearChangeD, "#.##%")
            Cells(counter1, 12).Value = Format(sumLL, "$#,###")
            
            If temp1 > 0 Then
                Cells(counter1, 10).Font.ColorIndex = 4
                Cells(counter1, 11).Interior.ColorIndex = 4
            Else
                Cells(counter1, 11).Interior.ColorIndex = 3
            End If
            
            ' increment + reassign values
            counter1 = counter1 + 1
            startYearD = Cells(i + 1, 3).Value
            yearChangeD = 0
            sumLL = 0
            counter2 = 0
            
        Else ' if old ticker
        
            counter2 = counter2 + 1
            sumLL = sumLL + Cells(i, 7).Value
        
        End If

    Next i

Next WS

End Sub


