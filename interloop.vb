Sub test2()

    For colNum = 1 To 10
        For rowNum = 1 To 100
            result = result + 1
            Cells(rowNum, colNum) = result
            
            If result > 50 Then
                With Cells(rowNum, colNum).Font
                .Color = RGB(255, 0, 0)
                .Bold = True
                End With
            End If
        
        Next rowNum
    Next colNum
End Sub



