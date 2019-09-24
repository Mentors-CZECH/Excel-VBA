Option Explicit
Dim radek_pamet As Long

Private Sub Worksheet_Activate()
    With ActiveCell.EntireRow.Font
        .Bold = False
        .Size = "10"
    End With
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    With ActiveCell.EntireRow.Font
        .Bold = False
        .Size = "10"
    End With
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    With Application
        .ScreenUpdating = False
    End With

    If radek_pamet > 0 Then
        With ActiveSheet.Cells(radek_pamet, 1)
            .EntireRow.Font.Bold = False
            .EntireRow.Font.Size = "10"
        End With
        Range("B" & radek_pamet & ":P" & radek_pamet).Interior.Color = xlNone
        
    End If

    If ActiveCell.Row >= 4 Then
    
        With ActiveCell.EntireRow.Font
            .Bold = True
            .Size = "12"
        End With
        radek_pamet = ActiveCell.Row
        Range("B" & radek_pamet & ":P" & radek_pamet).Interior.ColorIndex = 15
        ActiveCell.EntireColumn.AutoFit
    End If
    
    With Application
        .ScreenUpdating = True
    End With
End Sub

