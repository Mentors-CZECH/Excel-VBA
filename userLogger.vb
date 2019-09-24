Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
Dim RowLast As Integer
Dim StrFolderPath As String

    With Application
        .ScreenUpdating = False
    End With
    
    RowLast = ThisWorkbook.Worksheets("Users").Cells(Rows.Count, "A").End(xlUp).Row
    With ThisWorkbook.Worksheets("Users")
        .Range("A" & RowLast + 1).Value = Now()
        .Range("B" & RowLast + 1).Value = (Environ$("Username"))
    End With
    
    With Application
        .ScreenUpdating = True
    End With

End Sub

