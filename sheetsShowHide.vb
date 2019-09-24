Sub ZOBRAZ_LISTY()

Dim backendArray As Variant
Dim item As Variant

    backendArray = Array("sourceData", "Vzorce", "Ukončení_Manažeři", "PrimárníOkresy", "SekundárníOkresy", "PSČ-Okres_Data")
    With Application
    .ScreenUpdating = False
    End With
    
    If Sheets(backendArray(1)).Visible = 2 Then
        For Each item In backendArray
        Sheets(item).Visible = -1
        Next item
    Else
        For Each item In backendArray
        Sheets(item).Visible = 2
        Next item
    End If
    
    With Application
    .ScreenUpdating = True
    End With

End Sub
