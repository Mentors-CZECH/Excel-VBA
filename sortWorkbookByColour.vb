Sub SortWorkBookByColor()

Dim xArray1() As Long
Dim xArray2() As String
Dim n As Integer
Application.ScreenUpdating = False
If Val(Application.Version) >= 10 Then
    For i = 1 To Application.ActiveWorkbook.Worksheets.Count
        If Application.ActiveWorkbook.Worksheets(i).Visible = -1 Then
            n = n + 1
            ReDim Preserve xArray1(1 To n)
            ReDim Preserve xArray2(1 To n)
            xArray1(n) = Application.ActiveWorkbook.Worksheets(i).Tab.Color
            xArray2(n) = Application.ActiveWorkbook.Worksheets(i).Name
        End If
    Next
    For i = 1 To n
        For j = i To n
            If xArray1(j) < xArray1(i) Then
                temp = xArray2(i)
                xArray2(i) = xArray2(j)
                xArray2(j) = temp
                temp = xArray1(i)
                xArray1(i) = xArray1(j)
                xArray1(j) = temp
            End If
        Next
    Next
    For i = n To 1 Step -1
        Application.ActiveWorkbook.Worksheets(CStr(xArray2(i))).Move after:=Application.ActiveWorkbook.Worksheets(Application.ActiveWorkbook.Worksheets.Count)
    Next
End If
Application.ScreenUpdating = True
End Sub